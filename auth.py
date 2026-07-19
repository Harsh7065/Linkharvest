"""
auth.py
Sign-in gate backed by a Google Sheet, via a Google Apps Script Web App
(see apps_script.gs). The sheet itself is never exposed to the client -
only a true/false answer comes back over HTTPS.

Two things this module does:
1. verify_credentials() - the login-screen check, run once at startup.
2. start_periodic_recheck() - keeps re-validating in the background
   while the app is running, so revoking a user in the sheet takes
   effect within RECHECK_INTERVAL_SECONDS even mid-session, not just
   on next launch.

Network-failure policy (read this before wiring it up):
A network error is NOT treated as "invalid" - it raises AuthError
instead of silently blocking a legitimate user whose wifi hiccuped.
Only an explicit {"valid": false} response triggers a force-quit.
If you want strict "no internet = no access" instead, see the note
in verify_credentials().
"""
import hashlib
import threading
import time

import requests

AUTH_ENDPOINT = "Https://docs.google.com/spreadsheets/u/0/d/1qYYHGu_EXE_HUMyGacnsBbhQKPsDOuohtZ75ZTy2Kx0/htmlview"
TIMEOUT = 8
RECHECK_INTERVAL_SECONDS = 120  # how often to re-validate while the app runs


class AuthError(Exception):
    """Network/server problem reaching the auth backend - NOT the same as invalid credentials."""


def hash_password(password: str) -> str:
    """SHA-256 the password before it ever leaves the machine. Paste this
    module's output for a test password into the sheet's PasswordHash
    column - never paste a plaintext password there."""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verify_credentials(user_id: str, password: str) -> bool:
    """
    Returns True if id/password match a row in the linked sheet.
    Raises AuthError on network/server failure - the caller decides how
    to handle that (e.g. show "can't reach license server, check your
    connection" rather than treating it as a failed login).

    To go stricter ("must have internet, no offline grace at all"),
    have the caller treat AuthError the same as invalid credentials -
    that's a one-line change where you catch it in your login screen.
    """
    if not user_id or not user_id.strip() or not password:
        return False
    try:
        resp = requests.post(
            AUTH_ENDPOINT,
            json={"id": user_id.strip(), "password_hash": hash_password(password)},
            timeout=TIMEOUT,
        )
        resp.raise_for_status()
        data = resp.json()
        return bool(data.get("valid"))
    except requests.RequestException as e:
        raise AuthError(f"Could not reach the license server: {e}")
    except ValueError:
        raise AuthError("License server returned an unexpected response.")


def start_periodic_recheck(user_id: str, password: str, on_revoked, on_error=None):
    """
    Runs verify_credentials() every RECHECK_INTERVAL_SECONDS in a daemon
    thread for as long as the app is open.

    on_revoked: called with no args the moment a check comes back
        {"valid": false} - wire this to force-close the app (see the
        app.py integration note below).
    on_error: optional, called with the AuthError on network failure.
        Left unhandled by default (see network-failure policy above) -
        a dropped connection does NOT call on_revoked.

    Returns the Thread object (already started) in case you want to
    track/join it, though as a daemon thread it dies with the app.
    """
    def _loop():
        while True:
            time.sleep(RECHECK_INTERVAL_SECONDS)
            try:
                if not verify_credentials(user_id, password):
                    on_revoked()
                    return  # stop looping once the app is being killed
            except AuthError as e:
                if on_error:
                    on_error(e)
                # network hiccup: don't kill the session, just try again
                # next interval (see module docstring)

    t = threading.Thread(target=_loop, daemon=True)
    t.start()
    return t
