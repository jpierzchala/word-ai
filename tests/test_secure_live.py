import concurrent.futures
import http.client
import json
from pathlib import Path
import tempfile
import threading
import unittest
from http.server import ThreadingHTTPServer
from word_ai_mcp.secure_live import State, handler, rpc
from word_ai_mcp.document_state import GRANT

PARAGRAPH = "11111111-2222-4333-8444-555555555555"

class Clock:
    value = 100.0
    def __call__(self): return self.value

class StateTests(unittest.TestCase):
    def setUp(self):
        self.clock = Clock()
        self.state = State(clock=self.clock)
        self.url = "https://example.test/pilot.docx"
        self.connect()

    def connect(self):
        self.pair = self.state.office("pair", {"document": self.url, "access": GRANT})
        accepted = self.state.call("word_session_connect", {"pairing_id": self.pair["pairing_id"], "document": self.url})
        claimed = self.state.office("pair-status", {"pairing_id": self.pair["pairing_id"], "document": self.url}, self.pair["secret"])
        self.sid, self.cap = accepted["session_id"], claimed["capability"]
        c = self.queue("snapshot", scope="body", start=0, limit=50, query="")
        self.finish(c, {"ok": True, "paragraphs": [{"paragraph_id": PARAGRAPH, "text": "before"}]})

    def office(self, route, **body):
        return self.state.office(route, {"session_id": self.sid, "document": self.url, **body}, self.cap)

    def queue(self, name="read", **args):
        if name == "read": args = {"paragraph_id": PARAGRAPH, "start": 0, "limit": 20000, **args}
        return self.state.call("word_session_" + name, {"session_id": self.sid, **args})

    def finish(self, command, result):
        self.office("poll")
        self.office("begin", command_id=command["command_id"])
        return self.office("result", command_id=command["command_id"], result=result)

    def preview(self):
        c = self.queue("preview", paragraph_id=PARAGRAPH, operation="replace_paragraph", text="after", find="", expected_sha256="a"*64)
        return self.finish(c, {"ok": True, "text": "before", "ooxml_sha256": "a"*64})["result"]["preview_id"]

    def test_click_does_not_grant_access_until_local_mcp_connects(self):
        pair = self.state.office("pair", {"document": "second", "access": GRANT})
        response = self.state.office("pair-status", {"pairing_id": pair["pairing_id"], "document": "second"}, pair["secret"])
        self.assertEqual(response, {"status": "pending"})
        with self.assertRaises(ValueError):
            self.state.call("word_session_connect", {"pairing_id": pair["pairing_id"], "document": "wrong"})

    def test_document_consent_must_be_explicit(self):
        for access in [None, "selected-controls", "read-only"]:
            with self.assertRaises(ValueError): self.state.office("pair", {"document": self.url, "access": access})

    def test_pairing_secret_is_required_and_never_leaked_to_mcp(self):
        listed = json.dumps(self.state.call("word_session_list", {}))
        self.assertNotIn(self.cap, listed)
        self.assertNotIn(self.pair["secret"], listed)
        self.assertNotIn("before", listed)
        with self.assertRaises(PermissionError):
            self.state.office("pair-status", {"pairing_id": self.pair["pairing_id"], "document": self.url}, "wrong")

    def test_no_default_document_and_no_unindexed_paragraph(self):
        with self.assertRaises(ValueError): self.state.call("word_session_read", {"paragraph_id": PARAGRAPH})
        with self.assertRaises(ValueError): self.queue(paragraph_id="other")

    def test_document_and_capability_cannot_be_rebound(self):
        for body, token in [({"document": "different"}, self.cap), ({}, "wrong")]:
            with self.assertRaises(PermissionError):
                self.state.office("poll", {"session_id": self.sid, "document": self.url, **body}, token)

    def test_cancelled_write_cannot_begin(self):
        c = self.queue("apply", preview_id=self.preview(), request_id="edit-1")
        self.office("poll")
        self.queue("cancel", command_id=c["command_id"])
        with self.assertRaises(ValueError): self.office("begin", command_id=c["command_id"])

    def test_revocation_invalidates_dispatched_commands(self):
        c = self.queue("apply", preview_id=self.preview(), request_id="edit-1")
        self.office("poll"); self.office("disconnect")
        with self.assertRaises(ValueError): self.office("begin", command_id=c["command_id"])
        self.assertEqual((self.state.sessions, self.state.commands, self.state.previews), ({}, {}, {}))

    def test_pair_cancel_revokes_even_after_connect_race(self):
        self.state.office("pair-cancel", {"pairing_id": self.pair["pairing_id"], "document": self.url}, self.pair["secret"])
        self.assertFalse(self.state.sessions)

    def test_expiry_prevents_late_poll_and_begin(self):
        c = self.queue(); self.clock.value += 91
        self.assertIsNone(self.office("poll")["command"])
        self.assertEqual(self.queue("command_status", command_id=c["command_id"])["status"], "expired")
        with self.assertRaises(ValueError): self.office("begin", command_id=c["command_id"])

    def test_execution_timeout_is_unknown_and_blocks_retry(self):
        c = self.queue("apply", preview_id=self.preview(), request_id="edit-1")
        self.office("poll"); self.office("begin", command_id=c["command_id"]); self.clock.value += 16
        self.assertEqual(self.queue("command_status", command_id=c["command_id"])["status"], "unknown")
        with self.assertRaises(ValueError): self.queue()
        with self.assertRaises(ValueError): self.queue("cancel", command_id=c["command_id"])

    def test_read_timeout_does_not_create_unknown_write(self):
        c = self.queue(); self.office("poll"); self.office("begin", command_id=c["command_id"]); self.clock.value += 16
        self.assertEqual(self.queue("command_status", command_id=c["command_id"])["status"], "expired")
        self.queue()

    def test_request_id_is_idempotent_and_preview_one_use(self):
        pid = self.preview(); c = self.queue("apply", preview_id=pid, request_id="edit-1")
        self.assertEqual(c["command_id"], self.queue("apply", preview_id=pid, request_id="edit-1")["command_id"])
        with self.assertRaises(ValueError): self.queue("apply", preview_id=pid, request_id="edit-2")
        with self.assertRaises(ValueError): self.queue("apply", preview_id="other", request_id="edit-1")

    def test_only_one_consumer_can_claim_and_begin(self):
        c = self.queue()
        with concurrent.futures.ThreadPoolExecutor(8) as pool:
            results = list(pool.map(lambda _: self.office("poll"), range(8)))
        self.assertEqual(sum(r["command"] is not None for r in results), 1)
        self.office("begin", command_id=c["command_id"])
        with self.assertRaises(ValueError): self.office("begin", command_id=c["command_id"])

    def test_cross_session_preview_rejected(self):
        pid = self.preview(); self.connect()
        with self.assertRaises(ValueError): self.queue("apply", preview_id=pid, request_id="edit-1")

    def test_stale_preview_cannot_be_promoted_to_apply(self):
        c = self.queue("preview", paragraph_id=PARAGRAPH, operation="replace_paragraph", text="after", find="", expected_sha256="a"*64)
        r = self.finish(c, {"ok": True, "ooxml_sha256": "b"*64})
        self.assertEqual(r["status"], "failed"); self.assertFalse(self.state.previews)

    def test_retention_and_restart_revoke_access(self):
        self.preview(); self.clock.value += 901; self.state.clean()
        self.assertEqual((self.state.sessions, self.state.commands, self.state.previews, self.state.pairings), ({}, {}, {}, {}))
        with self.assertRaises(ValueError): State().call("word_session_read", {"session_id": self.sid, "paragraph_id": PARAGRAPH, "start": 0, "limit": 100})

    def test_no_legacy_tools_and_no_multiline_or_unbounded_requests(self):
        for text in ["two\nlines", "x"*20001]:
            with self.assertRaises(ValueError): self.queue("preview", paragraph_id=PARAGRAPH, operation="replace_paragraph", text=text, find="", expected_sha256="a"*64)
        with self.assertRaises(ValueError): self.queue("snapshot", scope="body", start=0, limit=1000, query="")
        self.assertTrue(rpc(self.state, {"id": 1, "method": "tools/call", "params": {"name": "docx_apply_patchset"}})["result"]["isError"])


class HttpTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), handler(State(), Path(cls.temp.name)))
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()
    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown(); cls.server.server_close(); cls.temp.cleanup()
    def request(self, path, body=None, headers=None, method="POST"):
        client = http.client.HTTPConnection("127.0.0.1", self.server.server_port)
        client.request(method, path, body=json.dumps(body or {}), headers={"Host": "localhost:3100", "Content-Type": "application/json", **(headers or {})})
        r = client.getresponse(); value = (r.status, r.read(), dict(r.getheaders())); client.close(); return value
    def test_all_data_endpoints_require_capability(self):
        for route in ["poll", "begin", "result", "disconnect", "pair-status", "pair-cancel"]:
            self.assertEqual(self.request("/office/" + route)[0], 401)
    def test_pairing_is_same_origin_metadata_only(self):
        body = {"document": "test.docx", "access": GRANT}
        self.assertEqual(self.request("/office/pair", body)[0], 403)
        self.assertEqual(self.request("/office/pair", body, {"Origin": "null"})[0], 403)
        status, data, _ = self.request("/office/pair", body, {"Origin": "https://localhost:3100"})
        self.assertEqual(status, 200); self.assertIn("pairing_id", json.loads(data)); self.assertNotIn("session_id", json.loads(data))
    def test_no_http_mcp_proxy_or_legacy_api(self):
        for path in ["/mcp", "/bridge//evil.test", "/office/register", "/office/read", "/../secrets/pairing-token"]:
            self.assertEqual(self.request(path)[0], 404)
    def test_host_and_origin_rejected(self):
        for h in [{"Origin": "null"}, {"Origin": "https://localhost:3101"}, {"Host": "evil.test"}]:
            self.assertEqual(self.request("/health", headers=h, method="GET")[0], 403)
    def test_office_host_query_and_health(self):
        Path(self.temp.name, "taskpane.html").write_text("panel")
        self.assertEqual(self.request("/taskpane.html?_host_Info=Word$Win32$16.01", method="GET")[0], 200)
        code, data, h = self.request("/health", method="GET")
        self.assertEqual(code, 200); self.assertEqual(json.loads(data), {"ok": True, "profile": "secure-live"})
        self.assertEqual(h["Cache-Control"], "no-store")

if __name__ == "__main__": unittest.main()
