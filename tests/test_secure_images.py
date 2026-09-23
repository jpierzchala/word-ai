import base64
import importlib.util
import json
from pathlib import Path
import struct
import sys
import tempfile
import unittest
import zlib
from word_ai_mcp.image_assets import image_info, decode_image, MAX_IMAGE_BYTES
from word_ai_mcp.document_state import State, GRANT, TOOLS
from word_ai_mcp.secure_live import rpc

spec = importlib.util.spec_from_file_location('image_adapter', Path(__file__).resolve().parents[1] / 'scripts/image_mcp_adapter.py')
adapter = importlib.util.module_from_spec(spec); spec.loader.exec_module(adapter)


def png(width=2, height=1):
    def chunk(name, data):
        return struct.pack('>I', len(data)) + name + data + struct.pack('>I', zlib.crc32(name + data))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\x00\xff\x00\x00\x00\xff\x00')) + chunk(b'IEND', b'')


class ImageTests(unittest.TestCase):
    def test_png_jpeg_dimensions_and_hash(self):
        self.assertEqual(image_info(png())['pixel_width'], 2)
        jpeg = b'\xff\xd8\xff\xc0\x00\x0b\x08\x00\x14\x00\x28\x01\x01\x11\x00\xff\xd9'
        self.assertEqual(image_info(jpeg)['pixel_height'], 20)
        self.assertEqual(decode_image(base64.b64encode(png()).decode())['bytes'], len(png()))

    def test_rejects_nonimages_malformed_encoding_large_and_pixel_bombs(self):
        for data in [b'<svg/>', b'secret', b'\xff\xd8\xff\xc0\x00\x01', png(8193, 1), png(8000, 8000), b'a'*(MAX_IMAGE_BYTES+1)]:
            with self.assertRaises(ValueError): image_info(data)
        with self.assertRaises(ValueError): decode_image('not base64!')

    def test_host_rejects_network_device_relative_and_nonimage_paths(self):
        for path in ['https://host/a.png', r'\\host\share\a.png', r'\\?\C:\a.png', 'C:a.png', r'C:\a.png:stream', r'C:\secrets.txt']:
            with self.assertRaises(ValueError): adapter.read_local_image(path)

    @unittest.skipUnless(sys.platform == 'win32', 'Windows local-path boundary')
    def test_host_reads_only_valid_local_image_and_hides_base64_from_tool_schema(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp)/'test.png'; path.write_bytes(png())
            encoded, name = adapter.read_local_image(str(path))
            self.assertEqual(base64.b64decode(encoded), png()); self.assertEqual(name, 'test.png')
            mapped = adapter.host_request({'method':'tools/call','params':{'name':adapter.REPLACE_IMAGE_TOOL,'arguments':{
                'session_id':'s','paragraph_id':'p','expected_sha256':'a'*64,'image_index':0,
                'image_path':str(path),'width_pt':0,'alt_text':'Replacement'}}})
            arguments = mapped['params']['arguments']
            self.assertNotIn('image_path', arguments); self.assertEqual(arguments['image_name'], 'test.png')
            self.assertEqual(base64.b64decode(arguments['image_base64']), png())
            path.write_bytes(b'not an image')
            with self.assertRaises(ValueError): adapter.read_local_image(str(path))
        result = adapter.host_response(rpc(State(), {'id': 1, 'method': 'tools/list'}), 'tools/list')
        image = next(t for t in result['result']['tools'] if t['name'] == adapter.IMAGE_TOOL)
        replacement = next(t for t in result['result']['tools'] if t['name'] == adapter.REPLACE_IMAGE_TOOL)
        self.assertIn('image_path', image['inputSchema']['required'])
        self.assertNotIn('image_base64', image['inputSchema']['properties'])
        self.assertIn('image_path', replacement['inputSchema']['required'])
        self.assertNotIn('image_base64', replacement['inputSchema']['properties'])

    def test_preview_apply_idempotency_does_not_return_image_bytes(self):
        state = State(); doc = 'synthetic.docx'
        ticket = state.office('pair', {'document': doc, 'access': GRANT})
        sid = state.call('word_session_connect', {'pairing_id': ticket['pairing_id'], 'document': doc})['session_id']
        claim = state.office('pair-status', {'pairing_id': ticket['pairing_id'], 'document': doc}, ticket['secret'])
        cap = claim['capability']; pid = '11111111-2222-4333-8444-555555555555'; state.sessions[sid]['paragraphs'].add(pid)
        encoded = base64.b64encode(png()).decode()
        args = dict(session_id=sid, paragraph_id=pid, expected_sha256='a'*64, image_base64=encoded, image_name='test.png', position='after', width_pt=144, caption='Caption', alt_text='Test image')
        command = state.call(adapter.IMAGE_TOOL, args)
        def office(route, **extra): return state.office(route, dict(document=doc,session_id=sid,**extra), cap)
        queued = office('poll')['command']; self.assertEqual(queued['payload']['height_pt'],72)
        office('begin',command_id=command['command_id'])
        preview = office('result',command_id=command['command_id'],result={'ok':True,'ooxml_sha256':'a'*64,'image_sha256':image_info(png())['sha256']})
        self.assertNotIn(encoded,json.dumps(preview))
        apply_args=dict(session_id=sid,preview_id=preview['result']['preview_id'],request_id='once')
        first=state.call('word_session_apply',apply_args); second=state.call('word_session_apply',apply_args)
        self.assertEqual(first['command_id'],second['command_id'])
        office('disconnect'); self.assertFalse(state.commands); self.assertFalse(state.previews)

    def test_image_preview_requires_discovered_anchor_and_valid_size(self):
        state=State(); state.sessions['s']={'seen':state.clock(),'paragraphs':{'p'},'requests':{}}
        args=dict(session_id='s',paragraph_id='p',expected_sha256='a'*64,image_base64=base64.b64encode(png()).decode(),image_name='test.png',position='after',width_pt=144,caption='',alt_text='')
        for changes in [{'width_pt':0},{'width_pt':501},{'caption':'two\nlines'},{'paragraph_id':'unknown'},{'position':'replace'},{'expected_sha256':'stale'}]:
            with self.assertRaises(ValueError):state.call(adapter.IMAGE_TOOL,{**args,**changes})

        replace_args=dict(session_id='s',paragraph_id='p',expected_sha256='a'*64,image_index=0,
                          image_base64=base64.b64encode(png()).decode(),image_name='test.png',width_pt=0,alt_text='Replacement')
        command=state.call(adapter.REPLACE_IMAGE_TOOL,replace_args)
        self.assertEqual(state.commands[command['command_id']]['payload']['operation'],'replace_image')
        state.commands[command['command_id']]['status']='succeeded'
        for changes in [{'width_pt':1},{'width_pt':501},{'image_index':1001},{'paragraph_id':'unknown'},{'expected_sha256':'stale'}]:
            with self.assertRaises(ValueError):state.call(adapter.REPLACE_IMAGE_TOOL,{**replace_args,**changes})

        delete_args=dict(session_id='s',paragraph_id='p',expected_sha256='a'*64,image_index=0)
        delete=state.call('word_session_preview_delete_image',delete_args)
        self.assertEqual(state.commands[delete['command_id']]['payload']['operation'],'delete_image')


if __name__ == '__main__': unittest.main()
