"""
app_stats.py
Tiny persisted counter + recent-activity log for the Dashboard home page.

Every feature page (Link Downloader, PDF Extractor, Data Profiler) calls one
of the record_* functions when it finishes a job. The Dashboard page reads
load_stats() each time it's shown, so the numbers and activity feed are
always current — no in-memory wiring needed between pages.

Stored at ~/.linkharvest/stats.json so counts survive app restarts.
"""
import os
import json
import datetime

STATS_DIR = os.path.join(os.path.expanduser("~"), ".linkharvest")
STATS_PATH = os.path.join(STATS_DIR, "stats.json")

MAX_ACTIVITY_ENTRIES = 20

_DEFAULTS = {
    "total_downloads": 0,
    "pdfs_processed": 0,
    "issues_resolved": 0,
    "activity": [],  # list of {"icon", "text", "detail", "timestamp"} newest-first
}


def load_stats():
    try:
        with open(STATS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        merged = dict(_DEFAULTS)
        merged.update(data)
        return merged
    except Exception:
        return dict(_DEFAULTS)


def _save_stats(stats):
    try:
        os.makedirs(STATS_DIR, exist_ok=True)
        with open(STATS_PATH, "w", encoding="utf-8") as f:
            json.dump(stats, f, indent=2)
    except Exception:
        pass  # stats are advisory-only; never let a write failure break a job


def _add_activity(stats, icon, text, detail):
    entry = {
        "icon": icon,
        "text": text,
        "detail": detail,
        "timestamp": datetime.datetime.now().isoformat(),
    }
    stats["activity"] = [entry] + stats.get("activity", [])[: MAX_ACTIVITY_ENTRIES - 1]


def record_download(ok_count, failed_count, folder_path):
    stats = load_stats()
    stats["total_downloads"] = stats.get("total_downloads", 0) + ok_count
    detail = os.path.basename(os.path.normpath(folder_path)) or folder_path
    _add_activity(stats, "downloader", f"Downloaded {ok_count} files", detail)
    _save_stats(stats)


def record_pdf_extraction(ok_count, failed_count, output_excel_path):
    stats = load_stats()
    stats["pdfs_processed"] = stats.get("pdfs_processed", 0) + ok_count
    detail = os.path.basename(output_excel_path)
    _add_activity(stats, "pdf", f"Extracted data from {ok_count} PDFs", detail)
    _save_stats(stats)


def record_profiler_clean(issues_resolved, save_path):
    stats = load_stats()
    stats["issues_resolved"] = stats.get("issues_resolved", 0) + max(issues_resolved, 0)
    detail = os.path.basename(save_path)
    _add_activity(stats, "profiler", "Data profiling completed", detail)
    _save_stats(stats)


def format_relative_time(iso_timestamp):
    """Mirrors the mockup's activity feed style: '10:24 AM' for today,
    'Yesterday' for yesterday, 'N days ago' further back."""
    try:
        ts = datetime.datetime.fromisoformat(iso_timestamp)
    except Exception:
        return ""
    now = datetime.datetime.now()
    delta_days = (now.date() - ts.date()).days
    if delta_days <= 0:
        return ts.strftime("%I:%M %p").lstrip("0")
    if delta_days == 1:
        return "Yesterday"
    return f"{delta_days} days ago"
