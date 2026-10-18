#!/usr/bin/env python3
"""D1 quota watchdog for StoreResolve.

Checks D1 reachability with a single cheap query (SELECT 1) via the
Cloudflare REST API — no Cloudflare Access session needed, ~1 row read,
so the check itself never meaningfully burns quota.

LIMITATION (learned 2026-10-18, the hard way): the REST API path and the
Worker's D1 binding do NOT share quota accounting. On 2026-10-18 this
script reported "healthy" while the Worker was genuinely blocked with
D1_QUOTA_EXCEEDED on every route. A REST-API exit 0 therefore means
"D1 API reachable", NOT "Worker healthy".

The authoritative Worker-side signal is GET /api/d1-ping (requires an
OWNER Access session) plus the frontend quota banner. This script is a
backup for total D1 outages only — never use its exit 0 to declare the
app healthy, and never run expensive D1 diagnostics from a watchdog.

Exit codes:
  0 = D1 API reachable (NOT a guarantee the Worker is unblocked)
  1 = D1 quota exceeded on the REST path (or D1 unreachable) -> alert the user
"""
import json
import sys
import urllib.request
import urllib.error

sys.path.insert(0, "/opt/muse/skills/custom-connector-setup/bin")
import dynamic_credentials as dc

ACCOUNT_ID = "5d41e9781d4b0a5d42e262ee441617e2"
DB_ID = "4f8132e3-3a81-4d66-9986-4656095e7779"


def main() -> int:
    entry = dc.dynamic_credential_entry("custom.cloudflare", "access_token")
    surrogate = entry["surrogate"]
    url = (
        f"https://api.cloudflare.com/client/v4/accounts/{ACCOUNT_ID}"
        f"/d1/database/{DB_ID}/query"
    )
    body = json.dumps({"sql": "SELECT 1 AS ok"}).encode()
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Authorization", "Bearer " + surrogate)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read())
    except urllib.error.HTTPError as e:
        print(f"D1_WATCHDOG: HTTP {e.code} reaching D1 API")
        return 1
    except Exception as e:  # noqa: BLE001 - watchdog must not crash silently
        print(f"D1_WATCHDOG: unreachable: {e}")
        return 1
    # Cloudflare returns 200 with success:false on quota errors
    if not data.get("success", True):
        errors = data.get("errors", [])
        msg = " ".join(e.get("message", "") for e in errors)
        print(f"D1_WATCHDOG: QUOTA_OR_ERROR: {msg[:300]}")
        return 1
    try:
        ok = data["result"][0]["results"][0]["ok"] == 1
    except (KeyError, IndexError, TypeError):
        print("D1_WATCHDOG: unexpected response shape")
        return 1
    if ok:
        print("D1_WATCHDOG: D1 API reachable (does NOT prove Worker unblocked)")
        return 0
    print("D1_WATCHDOG: unexpected ok value")
    return 1


if __name__ == "__main__":
    sys.exit(main())
