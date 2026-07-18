"""
updater.py
Checks GitHub's Releases API for a newer version than what's currently
running. Runs entirely over HTTPS to api.github.com, no auth needed
for public repos.

This does NOT silently replace the running .exe (a program can't
reliably overwrite its own running file on Windows). It just tells the
user a newer version exists and gives them a link to grab it.
"""
import requests

GITHUB_REPO = "Harsh7065/Linkharvest"  # "owner/repo"
API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
TIMEOUT = 6


def _parse_version(tag: str):
    """'v1.0.3' -> (1, 0, 3), for numeric comparison instead of string comparison."""
    cleaned = tag.lstrip("vV").strip()
    parts = []
    for p in cleaned.split("."):
        digits = "".join(c for c in p if c.isdigit())
        parts.append(int(digits) if digits else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def check_for_update(current_version: str):
    """
    Returns a dict {'version': 'v1.0.4', 'url': 'https://github.com/.../releases/tag/v1.0.4',
    'download_url': 'https://.../LinkHarvest.exe'} if a newer release exists,
    otherwise returns None. Returns None (fails silently) on any network error —
    an update check should never block or crash the app.
    """
    try:
        resp = requests.get(API_URL, timeout=TIMEOUT,
                             headers={"Accept": "application/vnd.github+json"})
        if resp.status_code != 200:
            return None
        data = resp.json()
        latest_tag = data.get("tag_name", "")
        if not latest_tag:
            return None

        if _parse_version(latest_tag) <= _parse_version(current_version):
            return None  # already up to date (or somehow ahead)

        download_url = None
        for asset in data.get("assets", []):
            if asset.get("name", "").lower().endswith(".exe"):
                download_url = asset.get("browser_download_url")
                break

        return {
            "version": latest_tag,
            "url": data.get("html_url", f"https://github.com/{GITHUB_REPO}/releases/latest"),
            "download_url": download_url,
        }
    except requests.RequestException:
        return None
    except (ValueError, KeyError):
        return None
