"""
HFR + EIS grid figures — split into Main and Other groups.
Row 0: HFR vs current density  (BP = black/gray, No BP = blue)
Row 1: EIS Nyquist              (BOL/30K/75K in black/gray)
Output: EIS/Claude/HFR_grid_Main.png, HFR_grid_Other.png
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.legend_handler import HandlerBase


class _TextOnlyHandler(HandlerBase):
    """Invisible marker — makes legend entry appear as text only."""
    def create_artists(self, legend, orig_handle,
                       x0, y0, width, height, fontsize, trans):
        from matplotlib.lines import Line2D
        return [Line2D([x0, x0], [y0, y0], linewidth=0, color="none", transform=trans)]


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
DURS = ["BOL", "30K", "75K"]

BP_COLORS  = {"BOL": "#000000", "30K": "#555555", "75K": "#aaaaaa"}
NBP_COLORS = {"BOL": "#1f77b4", "30K": "#6baed6", "75K": "#bdd7e7"}


def load_hfr(sheet_name):
    raw = pd.read_excel(XLFILE, sheet_name=sheet_name, header=None)
    data = {}
    for si, samp in enumerate(ALL_SAMPLES):
        base_col = si * 7
        entry = {}
        for di, dur in enumerate(DURS):
            ci  = raw.iloc[2:, base_col + di*2    ].astype(float).values
            hfr = raw.iloc[2:, base_col + di*2 + 1].astype(float).values
            mask = np.isfinite(ci) & np.isfinite(hfr)
            entry[dur] = (ci[mask], hfr[mask])
        data[samp] = entry
    return data


def load_eis():
    raw = pd.read_excel(XLFILE, sheet_name="EIS", header=None)
    data = {}
    for si, samp in enumerate(ALL_SAMPLES):
        base_col = si * 7
        entry = {}
        for di, dur in enumerate(DURS):
            zr = raw.iloc[2:, base_col + di*2    ].astype(float).values
            zi = raw.iloc[2:, base_col + di*2 + 1].astype(float).values
            mask = np.isfinite(zr) & np.isfinite(zi)
            entry[dur] = (zr[mask], zi[mask])
        data[samp] = entry
    return data


bp    = load_hfr("HFR BP")
no_bp = load_hfr("HFR no BP")
eis   = load_eis()


def make_figure(group_name, samples):
    fig, axes = plt.subplots(2, 4, figsize=(20, 12))

    # ── Row 0: HFR ──────────────────────────────────────────────────────────
    for ai, samp in enumerate(samples):
        ax = axes[0, ai]
        for dur in DURS:
            ci_bp, hfr_bp = bp[samp][dur]
            ci_no, hfr_no = no_bp[samp][dur]
            ax.plot(ci_bp, hfr_bp, color=BP_COLORS[dur],  lw=1.5)
            ax.plot(ci_no, hfr_no, color=NBP_COLORS[dur], lw=1.5)

        ax.set_title(LABELS[samp], fontsize=32, fontweight="bold", pad=14)
        ax.set_xlim(0, 2300)
        ax.set_ylim(0, 0.030)
        ax.xaxis.set_major_locator(ticker.MultipleLocator(1000))
        ax.xaxis.set_minor_locator(ticker.MultipleLocator(500))
        ax.yaxis.set_major_locator(ticker.MultipleLocator(0.01))
        ax.yaxis.set_minor_locator(ticker.MultipleLocator(0.005))
        ax.tick_params(which="major", labelsize=26, width=1.5, length=6)
        ax.tick_params(which="minor", width=1.0, length=3)

    axes[0, 0].set_ylabel("HFR (Ω)", fontsize=30)

    hdr_bp  = plt.Line2D([], [], color="none", label="BP")
    hdr_nbp = plt.Line2D([], [], color="none", label="No BP")
    col0 = [hdr_bp]  + [plt.Line2D([0],[0], color=BP_COLORS[d],  lw=1.5, label=d) for d in DURS]
    col1 = [hdr_nbp] + [plt.Line2D([0],[0], color=NBP_COLORS[d], lw=1.5, label=d) for d in DURS]
    leg0 = axes[0, 0].legend(
        handles=col0 + col1,
        handler_map={hdr_bp: _TextOnlyHandler(), hdr_nbp: _TextOnlyHandler()},
        loc="lower left", fontsize=20, frameon=False, ncol=2, handlelength=1.2, handleheight=0.7,
        columnspacing=0.4, labelspacing=0.2, borderpad=0.4,
    )
    for i, text in enumerate(leg0.get_texts()):
        if i in (0, 4):
            text.set_fontweight("bold")

    # ── Row 1: EIS Nyquist ──────────────────────────────────────────────────
    for ai, samp in enumerate(samples):
        ax = axes[1, ai]
        for dur in DURS:
            zr, zi_pos = eis[samp][dur]
            ax.plot(zr, zi_pos, color=BP_COLORS[dur], lw=1.0,
                    marker="o", markersize=10, markerfacecolor=BP_COLORS[dur])

        ax.set_xlim(0.030, 0.050)
        ax.set_ylim(0, 0.080)
        ax.xaxis.set_major_locator(ticker.MultipleLocator(0.01))
        ax.xaxis.set_minor_locator(ticker.MultipleLocator(0.005))
        ax.yaxis.set_major_locator(ticker.MultipleLocator(0.02))
        ax.yaxis.set_minor_locator(ticker.MultipleLocator(0.01))
        ax.tick_params(which="major", labelsize=26, width=1.5, length=6)
        ax.tick_params(which="minor", width=1.0, length=3)

    axes[1, 0].set_ylabel("Zimg (Ω)", fontsize=30)

    eis_handles = [plt.Line2D([0],[0], color=BP_COLORS[d], lw=1.5, label=d) for d in DURS]
    axes[1, 0].legend(
        handles=eis_handles, loc="upper left", fontsize=25,
        frameon=False,
        handlelength=1.4, labelspacing=0.3, borderpad=0.5,
    )

    plt.tight_layout(h_pad=4.0, rect=[0, 0.05, 1, 1])

    for row, label in enumerate(["Current Density (mA cm⁻²)", "Z' (Ω)"]):
        xc = (axes[row, 0].get_position().x0 + axes[row, -1].get_position().x1) / 2
        yb = axes[row, 0].get_position().y0 - 0.045
        fig.text(xc, yb, label, ha="center", va="top", fontsize=30)

    outfile = os.path.join(BASE, f"HFR_grid_{group_name}.png")
    fig.savefig(outfile, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved -> {outfile}")


make_figure("Main",  GROUPS["Main"])
make_figure("Other", GROUPS["Other"])
