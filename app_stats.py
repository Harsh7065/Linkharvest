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
    "total_failed_downloads": 0,
    "total_download_bytes": 0,
    "total_download_seconds": 0.0,
    "file_type_totals": {"images": 0, "videos": 0, "audio": 0, "documents": 0},
    "daily_downloads": {},  # {"YYYY-MM-DD": {"total": n, "completed": n}}
}
MAX_DAILY_ENTRIES = 60  # ~2 months of history is plenty for the "over time" chart


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


def record_download(ok_count, failed_count, folder_path, category_counts=None,
                     total_bytes=0, elapsed_seconds=0.0):
    stats = load_stats()
    stats["total_downloads"] = stats.get("total_downloads", 0) + ok_count
    stats["total_failed_downloads"] = stats.get("total_failed_downloads", 0) + max(failed_count, 0)
    stats["total_download_bytes"] = stats.get("total_download_bytes", 0) + max(total_bytes, 0)
    stats["total_download_seconds"] = stats.get("total_download_seconds", 0.0) + max(elapsed_seconds, 0.0)

    if category_counts:
        totals = stats.get("file_type_totals", dict(_DEFAULTS["file_type_totals"]))
        for key, count in category_counts.items():
            totals[key] = totals.get(key, 0) + count
        stats["file_type_totals"] = totals

    today = datetime.date.today().isoformat()
    daily = stats.get("daily_downloads", {})
    day_entry = daily.get(today, {"total": 0, "completed": 0})
    day_entry["total"] = day_entry.get("total", 0) + ok_count + failed_count
    day_entry["completed"] = day_entry.get("completed", 0) + ok_count
    daily[today] = day_entry
    # keep only the most recent MAX_DAILY_ENTRIES days so the file doesn't grow forever
    if len(daily) > MAX_DAILY_ENTRIES:
        for old_day in sorted(daily.keys())[: len(daily) - MAX_DAILY_ENTRIES]:
            del daily[old_day]
    stats["daily_downloads"] = daily

    detail = os.path.basename(os.path.normpath(folder_path)) or folder_path
    _add_activity(stats, "downloader", f"Downloaded {ok_count} files", detail)
    _save_stats(stats)


def get_analytics_summary(days=7):
    """Returns the aggregated numbers the Analytics page needs:
    per-day totals for the last `days` days (oldest-first), file-type
    totals, average speed (MB/s), success rate, total size, and total
    failed downloads."""
    stats = load_stats()

    daily = stats.get("daily_downloads", {})
    today = datetime.date.today()
    day_labels, day_totals, day_completed = [], [], []
    for offset in range(days - 1, -1, -1):
        d = today - datetime.timedelta(days=offset)
        key = d.isoformat()
        entry = daily.get(key, {"total": 0, "completed": 0})
        day_labels.append(d.strftime("%b %d"))
        day_totals.append(entry.get("total", 0))
        day_completed.append(entry.get("completed", 0))

    total_bytes = stats.get("total_download_bytes", 0)
    total_seconds = stats.get("total_download_seconds", 0.0)
    avg_speed_mbps = (total_bytes / (1024 * 1024)) / total_seconds if total_seconds > 0 else 0.0

    total_dl = stats.get("total_downloads", 0)
    total_failed = stats.get("total_failed_downloads", 0)
    attempted = total_dl + total_failed
    success_rate = (total_dl / attempted * 100) if attempted > 0 else 100.0

    return {
        "day_labels": day_labels,
        "day_totals": day_totals,
        "day_completed": day_completed,
        "file_type_totals": stats.get("file_type_totals", dict(_DEFAULTS["file_type_totals"])),
        "avg_speed_mbps": avg_speed_mbps,
        "success_rate": success_rate,
        "total_bytes": total_bytes,
        "total_failed": total_failed,
        "total_downloads": total_dl,
        "activity_totals": {
            "total_downloads": total_dl,
            "pdfs_processed": stats.get("pdfs_processed", 0),
            "issues_resolved": stats.get("issues_resolved", 0),
        },
    }


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
