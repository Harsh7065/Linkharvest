"""
analytics_chart.py
Charts for the Dashboard's "Download Analytics" page: a line chart of
downloads over the last N days, and a pie chart of file-type
distribution. Styled to match donut_chart.py's dark-UI palette so all
charts in the app look consistent.
"""

import matplotlib
matplotlib.use("Agg")  # safe default; the Tk page swaps in TkAgg via FigureCanvasTkAgg
import matplotlib.pyplot as plt

_BG = "#1F1F1F"
_TEXT = "#E6E6E6"
_GRID = "#3A3A3A"

_CATEGORY_COLORS = {
    "images": "#4DA3FF",
    "videos": "#FF7A45",
    "audio": "#8C8C8C",
    "documents": "#8b5cf6",
}
_CATEGORY_LABELS = {
    "images": "Images",
    "videos": "Videos",
    "audio": "Audio",
    "documents": "Documents",
}


def render_downloads_over_time(day_labels, day_totals, day_completed, figsize=(6.4, 3.2)):
    """Line chart: total vs completed downloads per day, oldest-first."""
    fig, ax = plt.subplots(figsize=figsize, facecolor=_BG)
    ax.set_facecolor(_BG)

    x = range(len(day_labels))
    has_data = any(day_totals) or any(day_completed)

    if not has_data:
        ax.text(0.5, 0.5, "No downloads yet — run the Link Downloader\nand this chart will fill in.",
                ha="center", va="center", color="#9A9A9A", fontsize=10,
                transform=ax.transAxes, linespacing=1.6)
        ax.set_xticks([])
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
    else:
        ax.plot(x, day_totals, marker="o", markersize=4, color="#d4a72c",
                 linewidth=2, label="Attempted")
        ax.plot(x, day_completed, marker="o", markersize=4, color="#e6e6e6",
                 linewidth=2, label="Completed")
        ax.set_xticks(list(x))
        ax.set_xticklabels(day_labels, color=_TEXT, fontsize=8, rotation=0)
        ax.tick_params(axis="y", colors=_TEXT, labelsize=8)
        ax.grid(True, axis="y", color=_GRID, linewidth=0.6, alpha=0.6)
        for spine in ax.spines.values():
            spine.set_color(_GRID)
        legend = ax.legend(loc="upper left", frameon=False, labelcolor=_TEXT, fontsize=8)

    fig.tight_layout()
    return fig


def render_activity_bar(activity_totals, figsize=(4.0, 3.2)):
    """Bar chart comparing Downloads / PDFs Extracted / Files Cleaned —
    all-time totals, one bright color per category for quick visual scanning."""
    labels = ["Downloads", "PDFs\nExtracted", "Files\nCleaned"]
    keys = ["total_downloads", "pdfs_processed", "issues_resolved"]
    colors = ["#2f6fed", "#2fae63", "#8b5cf6"]
    values = [activity_totals.get(k, 0) for k in keys]

    fig, ax = plt.subplots(figsize=figsize, facecolor=_BG)
    ax.set_facecolor(_BG)

    if not any(values):
        ax.text(0.5, 0.5, "No activity yet — use any feature\nand it'll show up here.",
                ha="center", va="center", color="#9A9A9A", fontsize=10,
                transform=ax.transAxes, linespacing=1.6)
        ax.set_xticks([])
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
    else:
        bars = ax.bar(labels, values, color=colors, width=0.55, edgecolor=_BG, linewidth=1)
        for rect, val in zip(bars, values):
            ax.text(rect.get_x() + rect.get_width() / 2, rect.get_height(),
                    f"{val:,}", ha="center", va="bottom", color=_TEXT, fontsize=9, fontweight="bold")
        ax.tick_params(axis="x", colors=_TEXT, labelsize=9)
        ax.tick_params(axis="y", colors=_TEXT, labelsize=8)
        ax.grid(True, axis="y", color=_GRID, linewidth=0.6, alpha=0.6)
        for spine in ax.spines.values():
            spine.set_color(_GRID)
        ax.set_ylim(0, max(values) * 1.2 if max(values) > 0 else 1)

    fig.tight_layout()
    return fig


def render_filetype_pie(file_type_totals, figsize=(4.0, 4.0)):
    """Pie chart of images/videos/audio/documents downloaded, all-time."""
    labels, sizes, colors = [], [], []
    for key in ("images", "videos", "audio", "documents"):
        count = file_type_totals.get(key, 0)
        if count <= 0:
            continue
        labels.append(f"{_CATEGORY_LABELS[key]} ({count})")
        sizes.append(count)
        colors.append(_CATEGORY_COLORS[key])

    fig, ax = plt.subplots(figsize=figsize, facecolor=_BG)
    ax.set_facecolor(_BG)

    if not sizes:
        ax.pie([1], colors=["#333333"], radius=1, wedgeprops=dict(edgecolor=_BG))
        ax.text(0, 0, "No files\ndownloaded yet", ha="center", va="center",
                color=_TEXT, fontsize=10, linespacing=1.5)
    else:
        wedges, _ = ax.pie(
            sizes, colors=colors, radius=1,
            wedgeprops=dict(edgecolor=_BG, linewidth=2),
            startangle=90,
        )
        ax.legend(
            wedges, labels,
            loc="center left",
            bbox_to_anchor=(1.0, 0.5),
            frameon=False,
            labelcolor=_TEXT,
            fontsize=8,
        )

    ax.set_aspect("equal")
    fig.tight_layout()
    return fig
