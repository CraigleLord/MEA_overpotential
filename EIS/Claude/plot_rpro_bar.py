"""
Grouped bar charts of proton transfer resistance (R_pro) from the Tafel tables.
One figure for Main samples, one for Other samples.
Output: EIS/Claude/Rpro_bar_Main.png, Rpro_bar_Other.png
"""

import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

BASE = os.path.dirname(os.path.abspath(__file__))

# R_pro values (Ω) read from Table_O2_15bp_*_alt.png
DATA = {
    "Main": {
        "CN-BM":     {"BOL": 0.0141, "30K": 0.0148, "75K": 0.0150},
        "KB-BM":     {"BOL": 0.0109, "30K": 0.0116, "75K": 0.0117},
        "VC-Polyol": {"BOL": 0.0038, "30K": 0.0029, "75K": 0.0028},
        "AB-Polyol": {"BOL": 0.0033, "30K": 0.0036, "75K": 0.0039},
    },
    "Other": {
        "CN-Polyol": {"BOL": 0.0130, "30K": 0.0140, "75K": 0.0117},
        "KB-Polyol": {"BOL": 0.0090, "30K": 0.0099, "75K": 0.0105},
        "VC-BM":     {"BOL": 0.0042, "30K": 0.0034, "75K": 0.0036},
        "AB-BM":     {"BOL": 0.0026, "30K": 0.0037, "75K": 0.0036},
    },
}

DURS       = ["BOL", "30K", "75K"]
DUR_COLORS = {"BOL": "#000000", "30K": "#555555", "75K": "#aaaaaa"}

# Sample colors: CN=dark green, KB=light blue, VC=black, AB=golden yellow
SAMPLE_COLORS = {
    "CN-BM":     "#2e7d32", "KB-BM":     "#87ceeb",
    "VC-Polyol": "#000000", "AB-Polyol": "#e6a817",
    "CN-Polyol": "#2e7d32", "KB-Polyol": "#87ceeb",
    "VC-BM":     "#000000", "AB-BM":     "#e6a817",
}

WIDTH  = 0.22   # individual bar width
OFFSETS = np.array([-1, 0, 1]) * WIDTH   # BOL, 30K, 75K offsets


def make_figure(group_name, data):
    samples = list(data.keys())
    x = np.arange(len(samples))

    fig, ax = plt.subplots(figsize=(9, 6))

    for di, dur in enumerate(DURS):
        vals = [data[s][dur] for s in samples]
        bars = ax.bar(x + OFFSETS[di], vals, WIDTH,
                      color=DUR_COLORS[dur], label=dur,
                      edgecolor="white", linewidth=0.5)

    ax.set_xticks(x)
    ax.set_xticklabels(samples, fontsize=22)
    ax.set_ylabel("R$_{pro}$ (Ω)", fontsize=26)
    ax.set_ylim(0, 0.018)
    ax.yaxis.set_major_locator(ticker.MultipleLocator(0.004))
    ax.yaxis.set_minor_locator(ticker.MultipleLocator(0.002))
    ax.tick_params(which="major", labelsize=22, width=1.5, length=6)
    ax.tick_params(which="minor", width=1.0, length=3)

    ax.legend(fontsize=20, frameon=False, handlelength=1.2,
              labelspacing=0.3, loc="upper right")

    # Color x-tick labels to match sample color convention
    for label, samp in zip(ax.get_xticklabels(), samples):
        label.set_color(SAMPLE_COLORS[samp])
        label.set_fontweight("bold")

    plt.tight_layout()
    outfile = os.path.join(BASE, f"Rpro_bar_{group_name}.png")
    fig.savefig(outfile, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved -> {outfile}")


make_figure("Main",  DATA["Main"])
make_figure("Other", DATA["Other"])
