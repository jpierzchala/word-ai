"""Exercise actual TLS + Docker STDIO using synthetic document-session state."""
import json
import os
from pathlib import Path
import socket
import ssl
import subprocess
import urllib.error
import urllib.request
context = ssl.create_default_context(cafile=str(Path(os.environ['LOCALAPPDATA']) / 'WordAiSecure/secrets/localhost.crt'))
def http(path, body=None, credential=None, headers=None):
    request = urllib.request.Request('https://localhost:3100' + path,
        data=json.dumps(body).encode() if body is not None else None,
        headers={'Content-Type': 'application/json', **({'Authorization': 'Bearer ' + credential} if credential else {}), **(headers or {})})
    with urllib.request.urlopen(request, context=context, timeout=5) as response:
        return json.load(response)
def rpc(method, params=None):
    request = {'jsonrpc': '2.0', 'id': 1, 'method': method, 'params': params or {}}
    proc = subprocess.run(['docker', 'exec', '-i', 'word-ai-secure', 'python', '/app/secure_live.py', 'stdio'],
        input=json.dumps(request) + '\n', text=True, capture_output=True, check=True, timeout=15)
    assert not proc.stderr, proc.stderr
    return json.loads(proc.stdout)['result']
def tool(name, args):
    response = rpc('tools/call', {'name': name, 'arguments': args})
    assert not response.get('isError'), response
    return json.loads(response['content'][0]['text'])
with socket.create_connection(('127.0.0.1', 3100), timeout=5):
    assert http('/health')['ok']
for path, expected, headers in [('/office/poll', 401, {}), ('/mcp', 404, {}),
    ('/office/pair', 403, {}), ('/office/pair', 403, {'Origin': 'null'}), ('/office/pair', 403, {'Host': 'evil.test'})]:
    try:
        http(path, {}, headers=headers)
        raise AssertionError('Unexpected access')
    except urllib.error.HTTPError as error:
        assert error.code == expected, (path, error.code)
assert rpc('initialize', {'protocolVersion': '2025-11-25'})['serverInfo']['version'] == '2.0.0'
listed_tools = {tool['name']: tool for tool in rpc('tools/list')['tools']}
assert len(listed_tools) == 13
assert {'word_session_preview_image', 'word_session_preview_replace_image', 'word_session_preview_delete_image',
        'word_session_preview_move_section', 'word_session_preview_table_row'} <= listed_tools.keys()
assert 'image_index' in listed_tools['word_session_preview_replace_image']['inputSchema']['required']
assert 'image_index' in listed_tools['word_session_preview_delete_image']['inputSchema']['required']
assert listed_tools['word_session_preview_table_row']['inputSchema']['properties']['cells']['type'] == 'array'
assert 'target_expected_sha256' in listed_tools['word_session_preview_move_section']['inputSchema']['required']
assert 'delete_paragraph' in listed_tools['word_session_preview']['description']
document = 'https://example.invalid/word-ai-synthetic-security-check.docx'
ticket = http('/office/pair', {'document': document, 'access': 'document-read-write-v2'}, headers={'Origin': 'https://localhost:3100'})
connection = tool('word_session_connect', {'pairing_id': ticket['pairing_id'], 'document': document})
registered = http('/office/pair-status', {'document': document, 'pairing_id': ticket['pairing_id']}, ticket['secret'])
sid, capability = registered['session_id'], registered['capability']
def office(route, **extra):
    return http('/office/' + route, {'document': document, 'session_id': sid, **extra}, capability)
def complete(queued, result):
    claimed = office('poll')['command']
    assert claimed['command_id'] == queued['command_id']
    office('begin', command_id=claimed['command_id'])
    return office('result', command_id=claimed['command_id'], result=result)['result']
try:
    listed = next(s for s in tool('word_session_list', {})['sessions'] if s['session_id'] == sid)
    assert set(listed) == {'session_id', 'document', 'access'}
    pid = '11111111-1111-1111-1111-111111111111'
    complete(tool('word_session_snapshot', {'session_id': sid, 'scope': 'body', 'start': 0, 'limit': 10, 'query': ''}), {'ok': True, 'paragraphs': [{'paragraph_id': pid}]})
    preview = complete(tool('word_session_preview', {'session_id': sid, 'paragraph_id': pid, 'operation': 'replace_text', 'find': 'before', 'text': 'after', 'expected_sha256': 'a'*64}), {'ok': True, 'before': 'before', 'after': 'after', 'ooxml_sha256': 'a'*64})
    # Structural previews are promoted only with the add-in's structural guard.
    target = '22222222-2222-2222-2222-222222222222'
    complete(tool('word_session_snapshot', {'session_id': sid, 'scope': 'body', 'start': 0, 'limit': 10, 'query': ''}), {'ok': True, 'paragraphs': [{'paragraph_id': pid}, {'paragraph_id': target}]})
    move = {'session_id': sid, 'paragraph_id': pid, 'expected_sha256': 'a'*64, 'target_paragraph_id': target, 'target_expected_sha256': 'b'*64, 'position': 'before'}
    unguarded = complete(tool('word_session_preview_move_section', move), {'ok': True, 'ooxml_sha256': 'a'*64, 'target_ooxml_sha256': 'b'*64})
    assert 'preview_id' not in unguarded, unguarded
    guarded = complete(tool('word_session_preview_move_section', move), {'ok': True, 'ooxml_sha256': 'a'*64, 'target_ooxml_sha256': 'b'*64, 'guard_sha256': 'c'*64})
    assert 'preview_id' in guarded, guarded
    args = {'session_id': sid, 'preview_id': preview['preview_id'], 'request_id': 'synthetic-once'}
    applied = tool('word_session_apply', args)
    assert tool('word_session_apply', args)['command_id'] == applied['command_id']
    office('poll')
    tool('word_session_cancel', {'session_id': sid, 'command_id': applied['command_id']})
    try:
        office('begin', command_id=applied['command_id'])
        raise AssertionError('Cancelled write started')
    except urllib.error.HTTPError as error:
        assert error.code == 400
finally:
    office('disconnect')
assert not any(s['session_id'] == sid for s in tool('word_session_list', {})['sessions'])
print('PASS: verified TLS, idle TLS peer, negative HTTP, thirteen MCP tools, two-sided pairing, snapshot discovery, preview, structural guard, idempotency, cancellation, revocation.')
