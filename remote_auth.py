"""
remote_auth.py
Optional remote "kill switch" / authorization gate. On startup, the app
fetches a small JSON control file you control (e.g. a raw file in your
GitHub repo) and decides whether to keep running.

Control file format (host this wherever you like, e.g.
https://raw.githubusercontent.com/<you>/<repo>/main/control.json):

{
  "active": true,
  "min_version": "1.2.0",
  "message": "This build has been retired. Please download the latest version."
}

- active: false           -> app shows `message` and exits immediately
- min_version higher than the running version -> same (forces an update)
- Any network/parse failure -> fails OPEN by default (app still runs).
  Set FAIL_CLOSED = True below if you'd rather block on failure (e.g. for
  paid/licensed software) instead of a free tool where you don't want a
  flaky connection to lock out a legit user.
"""
import time
import requests

CONTROL_URL = "https://raw.githubusercontent.com/Harsh7065/Linkharvest/main/control.json"
TIMEOUT = 5
FAIL_CLOSED = False  # True = block the app if the check itself fails


def _parse_version(v: str):
    cleaned = (v or "").lstrip("vV").strip()
    parts = []
    for p in cleaned.split("."):
        digits = "".join(c for c in p if c.isdigit())
        parts.append(int(digits) if digits else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def check_authorization(current_version: str):
    """
    Returns (allowed: bool, message: str | None).
    Call this once at startup, before the main window is built.
    """
    try:
        # cache-busting: GitHub's raw-file CDN caches responses at the edge
        # for several minutes, and different requests can even hit different
        # cache nodes. Appending a changing query param forces a fresh fetch
        # every time instead of risking a stale "active" value.
        bust = int(time.time())
        url = f"{CONTROL_URL}?_={bust}"
        resp = requests.get(url, timeout=TIMEOUT, headers={"Cache-Control": "no-cache"})
        if resp.status_code != 200:
            return (not FAIL_CLOSED, None)

        data = resp.json()

        if data.get("active", True) is False:
            return (False, data.get("message", "This app is no longer available."))

        min_version = data.get("min_version")
if min_version and _parse_version(current_version) < _parse_version(min_version):
    return (False, data.get(   # ← change False to True
        "message",
        f"A required update (v{min_version}+) is available. Please download the latest version."
           ))

        return (True, None)

    except requests.RequestException:
        return (not FAIL_CLOSED, None)
    except (ValueError, KeyError):
        return (not FAIL_CLOSED, None)
