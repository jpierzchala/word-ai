import json
import unittest
from word_ai_mcp.document_state import State, GRANT, TOOLS
from word_ai_mcp.secure_live import rpc

HEADING = "11111111-2222-4333-8444-555555555555"
TARGET = "66666666-2222-4333-8444-555555555555"


class Clock:
    value = 100.0
    def __call__(self): return self.value


class StructureTests(unittest.TestCase):
    def setUp(self):
        self.clock = Clock()
        self.state = State(clock=self.clock)
        self.url = "https://example.test/structure.docx"
        pair = self.state.office("pair", {"document": self.url, "access": GRANT})
        self.sid = self.state.call("word_session_connect", {"pairing_id": pair["pairing_id"], "document": self.url})["session_id"]
        self.cap = self.state.office("pair-status", {"pairing_id": pair["pairing_id"], "document": self.url}, pair["secret"])["capability"]
        c = self.call("snapshot", scope="body", start=0, limit=50, query="")
        self.finish(c, {"ok": True, "paragraphs": [{"paragraph_id": HEADING}, {"paragraph_id": TARGET}]})

    def call(self, name, **args):
        return self.state.call("word_session_" + name, {"session_id": self.sid, **args})

    def office(self, route, **body):
        return self.state.office(route, {"session_id": self.sid, "document": self.url, **body}, self.cap)

    def finish(self, command, result):
        self.office("poll")
        self.office("begin", command_id=command["command_id"])
        return self.office("result", command_id=command["command_id"], result=result)

    def move(self, **changes):
        return self.call("preview_move_section", **{"paragraph_id": HEADING, "expected_sha256": "a" * 64,
                         "target_paragraph_id": TARGET, "target_expected_sha256": "b" * 64, "position": "before", **changes})

    def edit(self, operation, text="", find=""):
        return self.call("preview", paragraph_id=HEADING, operation=operation, text=text, find=find, expected_sha256="a" * 64)

    def test_new_operations_validate_their_parameters(self):
        for operation, text, find in [("delete_paragraph", "", ""), ("format_text", "bold,no_italic", "Lead-in:"),
                                      ("format_text", "underline", ""), ("list_restart", "", ""), ("list_type", "bullet", ""),
                                      ("list_level", "8", "")]:
            c = self.edit(operation, text, find)
            self.state.commands[c["command_id"]]["status"] = "succeeded"
        for operation, text, find in [("delete_paragraph", "x", ""), ("delete_paragraph", "", "x"), ("format_text", "bold,no_bold", "x"),
                                      ("format_text", "heavy", "x"), ("format_text", "", "x"), ("format_text", "bold", "x" * 256),
                                      ("list_restart", "1", ""), ("list_type", "roman", ""), ("list_level", "9", ""),
                                      ("list_level", "1", "x"), ("format_text", "bold", "two\nlines")]:
            with self.assertRaises(ValueError, msg=(operation, text, find)):
                self.edit(operation, text, find)

    def test_move_requires_discovered_distinct_anchors_and_hashes(self):
        for changes in [{"target_paragraph_id": "undiscovered"}, {"paragraph_id": "undiscovered"}, {"target_paragraph_id": HEADING},
                        {"target_expected_sha256": "stale"}, {"expected_sha256": "B" * 64}, {"position": "inside"}]:
            with self.assertRaises(ValueError, msg=changes):
                self.move(**changes)
        c = self.move()
        self.assertEqual(self.state.commands[c["command_id"]]["payload"]["operation"], "move_section")

    def test_move_preview_checks_target_hash_and_carries_guard_into_apply(self):
        stale = self.finish(self.move(), {"ok": True, "ooxml_sha256": "a" * 64, "target_ooxml_sha256": "c" * 64, "guard_sha256": "d" * 64})
        self.assertEqual(stale["status"], "failed")
        missing = self.finish(self.move(), {"ok": True, "ooxml_sha256": "a" * 64, "target_ooxml_sha256": "b" * 64})
        self.assertEqual(missing["status"], "failed")
        preview = self.finish(self.move(), {"ok": True, "ooxml_sha256": "a" * 64, "target_ooxml_sha256": "b" * 64, "guard_sha256": "d" * 64})
        self.assertEqual(preview["status"], "succeeded")
        applied = self.call("apply", preview_id=preview["result"]["preview_id"], request_id="move-1")
        claimed = self.office("poll")["command"]
        self.assertEqual(claimed["command_id"], applied["command_id"])
        self.assertEqual(claimed["payload"]["guard_sha256"], "d" * 64)
        self.assertEqual(claimed["payload"]["target_expected_sha256"], "b" * 64)

    def test_structural_execution_gets_the_longer_deadline_and_text_writes_do_not(self):
        c = self.move()
        self.office("poll"); self.office("begin", command_id=c["command_id"])
        self.clock.value += 30
        self.assertEqual(self.call("command_status", command_id=c["command_id"])["status"], "executing")
        self.clock.value += 31
        self.assertEqual(self.call("command_status", command_id=c["command_id"])["status"], "expired")
        self.office("poll")  # the panel keeps polling while Word works
        d = self.edit("delete_paragraph")
        self.office("poll"); self.office("begin", command_id=d["command_id"])
        self.clock.value += 16
        self.assertEqual(self.call("command_status", command_id=d["command_id"])["status"], "executing")
        self.office("result", command_id=d["command_id"], result={"ok": False, "submitted": False})
        t = self.edit("replace_paragraph", "after")
        self.office("poll"); self.office("begin", command_id=t["command_id"])
        self.clock.value += 16
        self.assertEqual(self.call("command_status", command_id=t["command_id"])["status"], "expired")

    def test_guarded_edit_preview_without_guard_is_rejected(self):
        for operation in ["delete_paragraph", "list_restart"]:
            result = self.finish(self.edit(operation), {"ok": True, "ooxml_sha256": "a" * 64})
            self.assertEqual(result["status"], "failed", operation)
        result = self.finish(self.edit("format_text", "bold", "x"), {"ok": True, "ooxml_sha256": "a" * 64})
        self.assertEqual(result["status"], "succeeded")
        self.assertNotIn("guard_sha256", self.state.previews[result["result"]["preview_id"]]["payload"])

    def test_table_row_cells_are_bounded_single_line_strings(self):
        args = {"paragraph_id": HEADING, "expected_sha256": "a" * 64, "position": "after", "cells": ["Queue", ""]}
        c = self.call("preview_table_row", **args)
        self.assertEqual(self.state.commands[c["command_id"]]["payload"]["operation"], "insert_table_row")
        self.state.commands[c["command_id"]]["status"] = "succeeded"
        for cells in [[], "Queue", ["a\tb"], ["x" * 4001], [1], ["x"] * 64, ["x" * 4000] * 6]:
            with self.assertRaises(ValueError, msg=str(cells)[:40]):
                self.call("preview_table_row", **{**args, "cells": cells})
        with self.assertRaises(ValueError):
            self.call("preview_table_row", **{**args, "position": "inside"})

    def test_tool_schemas_expose_new_tools_with_array_cells(self):
        tools = {t["name"]: t for t in rpc(self.state, {"id": 1, "method": "tools/list"})["result"]["tools"]}
        self.assertEqual(len(tools), len(TOOLS))
        self.assertEqual(len(tools), 13)
        cells = tools["word_session_preview_table_row"]["inputSchema"]["properties"]["cells"]
        self.assertEqual(cells, {"type": "array", "items": {"type": "string"}, "minItems": 1, "maxItems": 63})
        self.assertTrue(tools["word_session_preview_move_section"]["annotations"]["readOnlyHint"])
        self.assertIn("delete_paragraph", tools["word_session_preview"]["description"])
        json.dumps(tools)


if __name__ == "__main__":
    unittest.main()
