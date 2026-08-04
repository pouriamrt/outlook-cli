"""One-shot probe: which scopes/audiences can the captured RT mint?"""

import base64
import json
import ssl
from pathlib import Path

import httpx
import truststore

creds = json.loads(Path.home().joinpath(".config/outlook-cli/credentials.json").read_text())
ctx = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)

scopes_to_try = [
    "https://outlook.office.com/.default",
    "https://outlook.office.com/Mail.ReadWrite offline_access",
    "https://graph.microsoft.com/Mail.Read offline_access",
    "Mail.Read offline_access",
]

for scope in scopes_to_try:
    r = httpx.post(
        f"https://login.microsoftonline.com/{creds['tenant_id']}/oauth2/v2.0/token",
        data={
            "client_id": creds["client_id"],
            "grant_type": "refresh_token",
            "refresh_token": creds["refresh_token"],
            "scope": scope,
        },
        headers={"Origin": "https://outlook.cloud.microsoft"},
        verify=ctx,
        timeout=30.0,
    )
    print(f"--- scope={scope!r} status={r.status_code}")
    try:
        body = r.json()
    except Exception:
        body = {"text": r.text[:200]}
    if "access_token" in body:
        p = body["access_token"].split(".")[1]
        p += "=" * (-len(p) % 4)
        claims = json.loads(base64.urlsafe_b64decode(p))
        print(f"  aud: {claims.get('aud')}")
        print(f"  scp: {claims.get('scp', '')[:300]}")
    else:
        print(f"  error: {body.get('error')}")
        print(f"  desc:  {body.get('error_description', '')[:300]}")
