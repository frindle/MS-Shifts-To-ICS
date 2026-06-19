#!/usr/bin/env python3
"""AMO unlisted signing script for Teams Shifts to ICS Firefox extension."""

import json, time, uuid, hashlib, hmac, base64, urllib.request, urllib.error, sys, os

def _load_env():
    env = {}
    env_path = os.path.join(os.path.dirname(__file__), '.env')
    if os.path.exists(env_path):
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    k, v = line.split('=', 1)
                    env[k.strip()] = v.strip()
    return env

_env = _load_env()
JWT_ISSUER = _env.get('AMO_JWT_ISSUER') or os.environ.get('AMO_JWT_ISSUER', '')
JWT_SECRET = _env.get('AMO_JWT_SECRET') or os.environ.get('AMO_JWT_SECRET', '')
if not JWT_ISSUER or not JWT_SECRET:
    print('ERROR: AMO_JWT_ISSUER and AMO_JWT_SECRET must be set in .env or environment')
    sys.exit(1)
ADDON_ID = '2993651'
XPI_PATH = '/Users/penndalton/Desktop/GitHub Projects/teams-shifts-exporter-firefox-1.63.xpi'
OUT_PATH = '/Users/penndalton/Desktop/GitHub Projects/teams-shifts-exporter-firefox-1.63-signed.xpi'
BASE = 'https://addons.mozilla.org/api/v5'


def b64url(data):
    return base64.urlsafe_b64encode(data).rstrip(b'=').decode()


def make_jwt():
    now = int(time.time())
    header = b64url(json.dumps({'alg': 'HS256', 'typ': 'JWT'}).encode())
    payload = b64url(json.dumps({'iss': JWT_ISSUER, 'iat': now, 'exp': now + 300, 'jti': str(uuid.uuid4())}).encode())
    sig_input = f'{header}.{payload}'.encode()
    sig = b64url(hmac.new(JWT_SECRET.encode(), sig_input, hashlib.sha256).digest())
    return f'{header}.{payload}.{sig}'


def api(method, path, data=None, files=None):
    url = f'{BASE}{path}'
    token = make_jwt()
    headers = {'Authorization': f'JWT {token}'}

    if files:
        boundary = uuid.uuid4().hex
        body_parts = []
        for name, (filename, content, ctype) in files.items():
            body_parts.append(
                f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"; filename="{filename}"\r\nContent-Type: {ctype}\r\n\r\n'.encode()
                + content + b'\r\n'
            )
        if data:
            for k, v in data.items():
                body_parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="{k}"\r\n\r\n{v}\r\n'.encode())
        body_parts.append(f'--{boundary}--\r\n'.encode())
        body = b''.join(body_parts)
        headers['Content-Type'] = f'multipart/form-data; boundary={boundary}'
        req = urllib.request.Request(url, data=body, headers=headers, method=method)
    elif data is not None:
        body = json.dumps(data).encode()
        headers['Content-Type'] = 'application/json'
        req = urllib.request.Request(url, data=body, headers=headers, method=method)
    else:
        req = urllib.request.Request(url, headers=headers, method=method)

    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as e:
        body = e.read()
        print(f'HTTP {e.code}: {body.decode()[:500]}')
        raise


def get_nested(obj, key):
    for part in key.split('.'):
        if not isinstance(obj, dict):
            return None
        obj = obj.get(part)
    return obj


def poll(path, key, target_values, interval=15, max_wait=600):
    deadline = time.time() + max_wait
    while time.time() < deadline:
        result = api('GET', path)
        val = get_nested(result, key)
        print(f'  {key}={val}')
        if val in target_values:
            return result
        time.sleep(interval)
    raise TimeoutError(f'Timed out waiting for {key} in {target_values}')


def download(url, dest):
    token = make_jwt()
    req = urllib.request.Request(url, headers={'Authorization': f'JWT {token}'})
    with urllib.request.urlopen(req) as r, open(dest, 'wb') as f:
        f.write(r.read())


TARGET_VERSION = '1.63'
print(f'=== AMO: Upload, sign, and download v{TARGET_VERSION} ===')

# 1. Upload XPI
print('\n[1] Uploading XPI...')
with open(XPI_PATH, 'rb') as f:
    xpi_bytes = f.read()
upload_result = api('POST', '/addons/upload/', files={
    'upload': ('extension.xpi', xpi_bytes, 'application/x-xpinstall'),
}, data={'channel': 'unlisted'})
upload_uuid = upload_result['uuid']
print(f'  uuid={upload_uuid}')

# 2. Poll until processed
print('\n[2] Waiting for upload validation...')
upload_info = poll(f'/addons/upload/{upload_uuid}/', 'processed', [True], interval=5, max_wait=300)
if not upload_info.get('valid'):
    print('Upload invalid:', json.dumps(upload_info.get('validation', {}), indent=2))
    sys.exit(1)
print('  Valid!')

# 3. Create version (may already exist if script was run before)
print('\n[3] Creating version...')
version_id = None
try:
    version_result = api('POST', f'/addons/addon/{ADDON_ID}/versions/', data={'upload': upload_uuid})
    version_id = version_result['id']
    print(f'  version_id={version_id}')
except urllib.error.HTTPError as e:
    if e.code == 409:
        print('  Version already exists, looking it up...')
        versions = api('GET', f'/addons/addon/{ADDON_ID}/versions/?filter=all_with_unlisted')
        for v in versions.get('results', []):
            if v.get('version') == TARGET_VERSION:
                version_id = v['id']
                print(f'  Found version_id={version_id}')
                break
        if not version_id:
            print('Could not find existing version')
            sys.exit(1)
    else:
        raise

# 4. Poll until signed
print('\n[4] Waiting for signing...')
version_info = poll(f'/addons/addon/{ADDON_ID}/versions/{version_id}/', 'file.status', ['public', 'disabled', 'failed'], interval=20, max_wait=600)
if version_info.get('file', {}).get('status') != 'public':
    print('Signing failed:', json.dumps(version_info, indent=2))
    sys.exit(1)

# 5. Download signed XPI
download_url = version_info['file']['url']
print(f'\n[5] Downloading signed XPI from {download_url}...')
download(download_url, OUT_PATH)
print(f'  Saved to {OUT_PATH}')
print('\nDone!')
sys.exit(0)

# --- fallback: find existing version ---
print(f'=== AMO: Find and download signed v{TARGET_VERSION} ===')

# Find version
print(f'\n[1] Looking up version {TARGET_VERSION}...')
versions = api('GET', f'/addons/addon/{ADDON_ID}/versions/?filter=all_with_unlisted')
version_info = None
for v in versions.get('results', []):
    if v.get('version') == TARGET_VERSION:
        version_id = v['id']
        version_info = v
        print(f'  Found version_id={version_id}, file status={v.get("file", {}).get("status")}')
        break

if not version_info:
    print('Version 1.19 not found. Full list:')
    for v in versions.get('results', []):
        print(f'  {v.get("version")} id={v["id"]} status={v.get("file", {}).get("status")}')
    sys.exit(1)

# Poll until signed
if version_info.get('file', {}).get('status') != 'public':
    print('\n[2] Waiting for signing...')
    version_info = poll(f'/addons/addon/{ADDON_ID}/versions/{version_id}/', 'file.status', ['public', 'disabled', 'failed'], interval=20, max_wait=600)

if version_info.get('file', {}).get('status') != 'public':
    print('Signing failed:', json.dumps(version_info, indent=2))
    sys.exit(1)

# Download signed XPI
download_url = version_info['file']['url']
print(f'\n[3] Downloading signed XPI from {download_url}...')
download(download_url, OUT_PATH)
print(f'  Saved to {OUT_PATH}')
print('\nDone!')
