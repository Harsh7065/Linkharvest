"""
donut_chart.py
Renders the anomalies-breakdown donut chart used by the Data Profiler page.
Kept separate from data_profiler.py so the plotting/UI dependency
(matplotlib) doesn't leak into pure backend code.
"""

import matplotlib
matplotlib.use("Agg")  # safe default; the Tk page swaps in TkAgg via FigureCanvasTkAgg
import matplotlib.pyplot as plt

from data_profiler import ANOMALY_LABELS

# Dark-UI-friendly palette, one color per anomaly type (order matches
# ANOMALY_LABELS insertion order so legend/colors stay stable).
_COLORS = {
    "missing_values": "#E8A33D",       # amber
    "duplicate_rows": "#FF7A45",       # neon orange
    "blank_rows": "#8C8C8C",           # slate gray
    "extra_spaces": "#4DA3FF",         # blue
    "special_characters": "#C77DFF",   # violet
    "mixed_data_types": "#FF5D8F",     # pink
}

_BG = "#1F1F1F"       # matches a typical dark customtkinter surface
_TEXT = "#E6E6E6"


def render_donut(results: dict, active_filters: dict, figsize=(4.2, 4.2)):
    """
    Build a matplotlib Figure showing a donut chart of anomaly counts.

    - `results` is the dict from analyze_data().
    - `active_filters` is {anomaly_key: bool} -- unchecked (False) items
      are excluded from the chart, matching "uncheck a box, recompute
      live" from the mockup.
    - Zero-count items are always excluded regardless of filter state.

    Returns the Figure; caller is responsible for embedding it
    (FigureCanvasTkAgg) or saving it.
    """
    labels, sizes, colors = [], [], []
    for key, label in ANOMALY_LABELS.items():
        if not active_filters.get(key, False):
            continue
        count = results.get(key, 0)
        if count <= 0:
            continue
        labels.append(f"{label}\n({count})")
        sizes.append(count)
        colors.append(_COLORS.get(key, "#999999"))

    fig, ax = plt.subplots(figsize=figsize, facecolor=_BG)
    ax.set_facecolor(_BG)

    if not sizes:
        # Empty state: nothing selected / nothing found.
        ax.pie([1], colors=["#333333"], radius=1, wedgeprops=dict(width=0.38, edgecolor=_BG))
        ax.text(0, 0, "No anomalies\nselected", ha="center", va="center",
                color=_TEXT, fontsize=11, linespacing=1.5)
    else:
        wedges, _ = ax.pie(
            sizes,
            colors=colors,
            radius=1,
            wedgeprops=dict(width=0.38, edgecolor=_BG, linewidth=2),
            startangle=90,
        )
        total = sum(sizes)
        ax.text(0, 0.08, f"{total}", ha="center", va="center",
                color=_TEXT, fontsize=20, fontweight="bold")
        ax.text(0, -0.14, "total anomalies", ha="center", va="center",
                color="#9A9A9A", fontsize=9)
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
