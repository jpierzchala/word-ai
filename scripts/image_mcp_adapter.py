"""Translate one local-image tool at the trusted host boundary; never log image bytes."""
import base64
import copy
import ctypes
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'word_ai_mcp'))
from image_assets import MAX_IMAGE_BYTES, image_info

IMAGE_TOOL = 'word_session_preview_image'
REPLACE_IMAGE_TOOL = 'word_session_preview_replace_image'
IMAGE_TOOLS = {IMAGE_TOOL, REPLACE_IMAGE_TOOL}
MAX_FRAME = 4 * 1024 * 1024


def read_local_image(value):
    # Deliberately excludes UNC/URLs, device paths, drive-relative paths and ADS.
    if not isinstance(value, str) or not re.fullmatch(r'[A-Za-z]:[\\/][^:\x00-\x1f]+', value):
        raise ValueError('Use an absolute local drive path to one PNG/JPEG file')
    path = Path(value)
    if sys.platform == 'win32' and ctypes.windll.kernel32.GetDriveTypeW(str(path.anchor)) not in {2, 3}:
        raise ValueError('Image must be on a local fixed or removable drive, not a mapped network drive')
    if path.suffix.lower() not in {'.png', '.jpg', '.jpeg'}:
        raise ValueError('Only .png, .jpg and .jpeg files are allowed')
    for part in [path, *path.parents]:
        info = part.lstat()
        if stat.S_ISLNK(info.st_mode) or getattr(info, 'st_file_attributes', 0) & 0x400:
            raise ValueError('Image path must not pass through symbolic links or reparse points')
    with path.open('rb') as stream:
        info = os.fstat(stream.fileno())
        if not stat.S_ISREG(info.st_mode) or info.st_size > MAX_IMAGE_BYTES:
            raise ValueError('Image must be a regular file of at most 2 MiB')
        data = stream.read(MAX_IMAGE_BYTES + 1)
    image_info(data)
    return base64.b64encode(data).decode('ascii'), path.name


def host_request(request):
    request = copy.deepcopy(request)
    params = request.get('params', {})
    if request.get('method') == 'tools/call' and params.get('name') in IMAGE_TOOLS:
        args = params.get('arguments', {})
        expected = ({'session_id', 'paragraph_id', 'expected_sha256', 'image_path', 'position', 'width_pt', 'caption', 'alt_text'}
                    if params.get('name') == IMAGE_TOOL else
                    {'session_id', 'paragraph_id', 'expected_sha256', 'image_index', 'image_path', 'width_pt', 'alt_text'})
        if not isinstance(args, dict) or set(args) != expected:
            raise ValueError('Use the documented image preview arguments, including image_path')
        encoded, name = read_local_image(args.pop('image_path'))
        args.update(image_base64=encoded, image_name=name)
    return request


def host_response(response, method):
    if method == 'tools/list':
        for tool in response.get('result', {}).get('tools', []):
            if tool['name'] in IMAGE_TOOLS:
                schema = tool['inputSchema']
                schema['properties'].pop('image_base64'); schema['properties'].pop('image_name')
                schema['properties']['image_path'] = {'type': 'string', 'description': 'Absolute local Windows path to the user-selected PNG/JPEG; not a URL or network path.'}
                schema['required'] = [x for x in schema['required'] if x not in {'image_base64', 'image_name'}] + ['image_path']
                tool['description'] = tool['description'].replace('image_base64 and image_name', 'image_path (local file; bytes are transferred outside model context)')
    return response


def relay(command):
    child = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                             stderr=sys.stderr, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    try:
        while raw := sys.stdin.buffer.readline(262145):
            if len(raw) > 262144: raise ValueError('MCP request too large')
            request = json.loads(raw)
            try:
                mapped = host_request(request)
            except (ValueError, OSError):
                # No file contents, secrets or exception payloads in model-visible diagnostics.
                response = {'jsonrpc': '2.0', 'id': request.get('id'), 'result': {'isError': True,
                    'content': [{'type': 'text', 'text': 'Image rejected: use an existing regular local PNG/JPEG, no links/UNC/URLs, <=2 MiB, <=8192 per side, <=20 MP; check arguments.'}]}}
            else:
                child.stdin.write(json.dumps(mapped).encode('utf-8') + b'\n'); child.stdin.flush()
                if 'id' not in request: continue
                line = child.stdout.readline(MAX_FRAME + 1)
                if not line or len(line) > MAX_FRAME: raise RuntimeError('MCP transport closed or response too large')
                response = host_response(json.loads(line), request.get('method'))
            sys.stdout.buffer.write(json.dumps(response, ensure_ascii=False).encode('utf-8') + b'\n'); sys.stdout.buffer.flush()
    finally:
        child.stdin.close()
        try: child.wait(timeout=5)
        except subprocess.TimeoutExpired: child.terminate(); child.wait(timeout=5)
    return child.returncode
