"""
Renders individual Tafel panels sized for the combined 4-column figure.
Large fonts, no title, no axis labels (those are added by combine_tafel_subplots.py).
Output: Tafel_large/ subfolder
"""

import sys, os, openpyxl
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy import stats

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

AREA  = 5.0
E_EQ  = 1.23
# Resolve paths relative to the original script location (two levels up)
HERE  = os.path.dirname(os.path.abspath(__file__))
BASE  = os.path.normpath(os.path.join(HERE, "..", ".."))
OUT   = os.path.join(HERE, "Tafel_large")
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "font.family":       "Arial",
    "font.size":         22,
    "axes.linewidth":    1.5,
    "xtick.major.width": 1.5, "ytick.major.width": 1.5,
    "xtick.minor.width": 1.0, "ytick.minor.width": 1.0,
    "xtick.direction":   "out", "ytick.direction": "out",
    "xtick.top":  False, "ytick.right": False,
    "xtick.major.size":  6,   "ytick.major.size": 6,
    "xtick.minor.size":  3.5, "ytick.minor.size": 3.5,
})

# (internal key, data filename, display label, show_legend)
# CN-BM and CN-Polyol are leftmost in their respective combined figures
SAMPLES = [
    ("CN BM",     "CN BM 1.5bp O2",     "CN-BM",     True),
    ("CN Polyol", "CN Polyol 1.5bp O2", "CN-Polyol", True),
    ("KB BM",     "KB BM 1.5bp O2",     "KB-BM",     False),
    ("KB Polyol", "KB Polyol 1.5bp O2", "KB-Polyol", False),
    ("VC BM",     "VC BM 1.5bp O2",     "VC-BM",     False),
    ("VC Polyol", "VC Polyol 1.5bp O2", "VC-Polyol", False),
    ("AB BM",     "AB BM 1.5bp O2",     "AB-BM",     False),
    ("AB Polyol", "AB Polyol 1.5bp O2", "AB-Polyol", False),
]

DURS = [
    ("BOL", "BOL",  "#000000", "red",   "-",  "s"),
    ("30K", "30K",  "#555555", "blue",  "--", "o"),
    ("75K", "75k",  "#999999", "green", ":",  "^"),
]

LN_LO, LN_HI = -3.5, -2.0

WINDOW_OVERRIDE = {
    ("CN BM",    "30K"): (-4.0, -1.7),
    ("VC Polyol","30K"): (-3.5, -1.9),
}


def load_iv(dur_dir, fname):
    folder = os.path.join(BASE, "Overpotnital", "O2 BP", dur_dir, "Edited")
    for sfx in ("_edited.xlsx", "_unchanged.xlsx"):
        p = os.path.join(folder, fname + sfx)
        if os.path.exists(p):
            wb = openpyxl.load_workbook(p, data_only=True)
            ws = wb.active
            pts = [(ws.cell(r, 1).value, ws.cell(r, 2).value)
                   for r in range(4, ws.max_row + 1)
                   if isinstance(ws.cell(r, 1).value, (int, float))
                   and isinstance(ws.cell(r, 2).value, (int, float))]
            wb.close()
            I_A = np.array([a for a, _ in pts])
            V   = np.array([v for _, v in pts])
            return I_A / AREA, V
    return None, None


def tafel_fit(j, eta, samp_lbl, dur_lbl):
    ln_j = np.log(j)
    lo, hi = WINDOW_OVERRIDE.get((samp_lbl, dur_lbl), (LN_LO, LN_HI))
    mask = (ln_j >= lo) & (ln_j <= hi)
    if mask.sum() < 2:
        return None, None, ln_j
    slope, intercept, *_ = stats.linregress(ln_j[mask], eta[mask])
    return slope, intercept, ln_j


for samp_lbl, fname, display_label, show_legend in SAMPLES:
    fig, ax = plt.subplots(figsize=(4.5, 4.0), dpi=300)

    all_eta = []

    for dur_lbl, dur_dir, mcolor, lcolor, lst, mkr in DURS:
        j, V = load_iv(dur_dir, fname)
        if j is None:
            print(f"{samp_lbl:<12} {dur_lbl:<4} MISSING")
            continue

        eta = V - E_EQ
        slope, intercept, ln_j = tafel_fit(j, eta, samp_lbl, dur_lbl)
        all_eta.extend(eta.tolist())

        ax.plot(ln_j, eta, marker=mkr, color=mcolor, markersize=7,
                linestyle="none", markeredgewidth=0, zorder=3, label=dur_lbl)

        if slope is not None:
            x_line = np.linspace(-5, 1, 300)
            ax.plot(x_line, slope * x_line + intercept,
                    color=lcolor, ls="--", lw=1.8, zorder=4)

    y_min = min(all_eta) if all_eta else -1.05
    y_min = np.floor(y_min * 10) / 10 - 0.05
    ax.set_xlim(-5, 1)
    ax.set_ylim(y_min, -0.30)

    # Title centered over the data box; no x/y axis labels (added by combine script)
    ax.set_title(display_label, fontsize=26, fontweight="bold", pad=8)
    ax.xaxis.set_major_locator(plt.MultipleLocator(2))
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.2))
    ax.tick_params(which="major", labelsize=26, width=1.5, length=6, direction="out",
                   top=False, right=False)
    ax.tick_params(which="minor", width=1.0, length=3.5, direction="out",
                   top=False, right=False)
    ax.minorticks_on()

    if show_legend:
        ax.legend(loc="lower left", fontsize=18, markerscale=1.8,
                  handlelength=0.8, borderpad=0.5, labelspacing=0.3,
                  frameon=False)

    plt.tight_layout()
    safe = samp_lbl.replace(" ", "")
    out  = os.path.join(OUT, f"Tafel_{safe}_O2BP_alt.png")
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {os.path.basename(out)}")
