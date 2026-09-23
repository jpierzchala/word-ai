"""Bounded PNG/JPEG header inspection. Decoding is checked separately in Word's webview."""
import base64
import binascii
import hashlib
import struct

MAX_IMAGE_BYTES = 2 * 1024 * 1024
MAX_IMAGE_BASE64 = ((MAX_IMAGE_BYTES + 2) // 3) * 4


def image_info(data):
    if not isinstance(data, bytes) or not 0 < len(data) <= MAX_IMAGE_BYTES:
        raise ValueError("Image must be PNG/JPEG, at most 2 MiB")
    width = height = 0
    if data.startswith(b'\x89PNG\r\n\x1a\n'):
        if len(data) < 33 or data[8:16] != b'\x00\x00\x00\rIHDR':
            raise ValueError("Invalid PNG header")
        width, height = struct.unpack('>II', data[16:24])
        mime = 'image/png'
    elif data.startswith(b'\xff\xd8'):
        mime = 'image/jpeg'
        pos = 2
        while pos + 4 <= len(data):
            if data[pos] != 255:
                raise ValueError("Invalid JPEG markers")
            while pos < len(data) and data[pos] == 255:
                pos += 1
            if pos >= len(data): break
            marker = data[pos]; pos += 1
            if marker in (0xD9, 0xDA): break
            if marker == 0x01 or 0xD0 <= marker <= 0xD7: continue
            size = int.from_bytes(data[pos:pos+2], 'big')
            if size < 2 or pos + size > len(data):
                raise ValueError("Invalid JPEG segment")
            if marker in (0xC0, 0xC1, 0xC2):
                if size < 8: raise ValueError("Invalid JPEG frame")
                height, width = struct.unpack('>HH', data[pos+3:pos+7])
                break
            pos += size
    else:
        raise ValueError("Only PNG and JPEG are supported; SVG, URLs and other files are rejected")
    if not 0 < width <= 8192 or not 0 < height <= 8192 or width * height > 20000000:
        raise ValueError("Image dimensions must be <=8192 per side and <=20 million pixels")
    return {'mime': mime, 'pixel_width': width, 'pixel_height': height,
            'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}


def decode_image(encoded):
    if not isinstance(encoded, str) or not 0 < len(encoded) <= MAX_IMAGE_BASE64:
        raise ValueError("Image exceeds the 2 MiB limit")
    try:
        data = base64.b64decode(encoded, validate=True)
    except (ValueError, binascii.Error) as error:
        raise ValueError("Invalid image encoding") from error
    return image_info(data)
