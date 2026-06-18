"""
EIS Nyquist — BOL only, split into Main and Other groups.
Single panel per figure, all 4 samples overlaid.
Output: EIS/Claude/EIS_BOL_Main.png, EIS_BOL_Other.png
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker


BASE   = os.path.dirname(os.path.abspath(__file__))
XLFILE = os.path.join(BASE, "Raw data EIS+HFR.xlsx")

ALL_SAMPLES = ["CN BM", "KB BM", "VC Polyol", "AB Polyol",
               "CN Polyol", "KB Polyol", "VC BM", "AB BM"]
GROUPS = {
    "Main":  ["CN BM", "KB BM", "VC Polyol", "AB Polyol"],
    "Other": ["CN Polyol", "KB Polyol", "VC BM", "AB BM"],
}
LABELS = {
    "CN BM":     "CN-BM",     "KB BM":     "KB-BM",
    "VC Polyol": "VC-Polyol", "AB Polyol": "AB-Polyol",
    "CN Polyol": "CN-Polyol", "KB Polyol": "KB-Polyol",
    "VC BM":     "VC-BM",     "AB BM":     "AB-BM",
}
# CN=dark green, KB=light blue, VC=black, AB=golden yellow  (matches LSV figure convention)
COLORS = ["#2e7d32", "#87ceeb", "#000000", "#e6a817"]


def load_eis():
    raw = pd.read_excel(XLFILE, sheet_name="EIS", header=None)
    data = {}
    for si, samp in enumerate(ALL_SAMPLES):
        base_col = si * 7
        zr = raw.iloc[2:, base_col].astype(float).values
        zi = raw.iloc[2:, base_col + 1].astype(float).values
        mask = np.isfinite(zr) & np.isfinite(zi)
        data[samp] = (zr[mask], zi[mask])
    return data


eis = load_eis()


def make_figure(group_name, samples):
    fig, ax = plt.subplots(figsize=(7, 7))

    for color, samp in zip(COLORS, samples):
        zr, zi = eis[samp]
        ax.plot(zr, zi, color=color, lw=1.5,
                marker="o", markersize=8, markerfacecolor=color,
                label=LABELS[samp])

    ax.set_xlim(0.030, 0.050)
    ax.set_ylim(0, 0.080)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(0.01))
    ax.xaxis.set_minor_locator(ticker.MultipleLocator(0.005))
    ax.yaxis.set_major_locator(ticker.MultipleLocator(0.02))
    ax.yaxis.set_minor_locator(ticker.MultipleLocator(0.01))
    ax.tick_params(which="major", labelsize=22, width=1.5, length=6)
    ax.tick_params(which="minor", width=1.0, length=3)
    ax.set_xlabel("Z' (Ω)", fontsize=26)
    ax.set_ylabel("Z'' (Ω)", fontsize=26)
    ax.legend(fontsize=20, frameon=False, handlelength=1.4, labelspacing=0.3)

    plt.tight_layout()

    outfile = os.path.join(BASE, f"EIS_BOL_{group_name}.png")
    fig.savefig(outfile, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved -> {outfile}")


make_figure("Main",  GROUPS["Main"])
make_figure("Other", GROUPS["Other"])
