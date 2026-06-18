"""
KB BM power-density panels (4 conditions) combined into ONE figure:
  4 rows x 1 column, LINEAR x-axis (A cm^-2, not log).
  No titles, no axis-title labels — tick numbers only.
  BOL (solid, diamond) + 75K (dashed, plus) per panel, shared y-scale.
Output -> 260603 New data set (reproducibility)/Final Plots/PD plots/
          IV_PD_KB_BM_AllCond_PD_linear.png
"""

import os, sys, openpyxl
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
from matplotlib.ticker import MaxNLocator, FuncFormatter, MultipleLocator

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE      = os.path.dirname(os.path.abspath(__file__))
AREA      = 5.0
REPRO_DIR = os.path.join(BASE, "260603 New data set (reproducibility)")
OUT_DIR   = os.path.join(REPRO_DIR, "Final Plots", "PD plots")
os.makedirs(OUT_DIR, exist_ok=True)
VC_POLYOL_OVERRIDE = os.path.join(REPRO_DIR, "Overpotnital")

# Per-condition data overrides: point a condition at the 260603 staging files.
# Staging layout differs: files sit directly in <dir>/<dur>/ (no "Edited" subfolder).
REPRO_ALT = os.path.join(REPRO_DIR, "Overpotential DATA raw alt")
OVERRIDE_DIRS = {
    "Air BP": os.path.join(REPRO_ALT, "Air BP", "KB BM"),
}

FS = 22
plt.rcParams.update({
    "font.family":       "Arial",
    "font.size":         FS,
    "axes.linewidth":    1.0,
    "xtick.major.width": 1.2, "ytick.major.width": 1.2,
    "xtick.minor.width": 0.8, "ytick.minor.width": 0.8,
    "xtick.direction":   "out", "ytick.direction": "out",
    "xtick.top":  False, "ytick.right": False,
    "xtick.major.size":  5,   "ytick.major.size":  5,
    "lines.linewidth":   1.6,
})

ROWS = [
    ("Air",    "75K"),
    ("Air BP", "75K"),
    ("O2",     "75K"),
    ("O2 BP",  "75k"),
]

FNAMES = {
    ("Air",    "KB BM"): "KB BM air",
    ("Air BP", "KB BM"): "KB BM 1.5bp air",
    ("O2",     "KB BM"): "KB BM O2",
    ("O2 BP",  "KB BM"): "KB BM 1.5bp O2",
}

DUR_META = {
    "BOL": dict(ls="-",  marker="D", color="#000000"),
    "75K": dict(ls="--", marker="P", color="#444444"),
}
DUR_ORDER = ["BOL", "75K"]

SAMPLE = "KB BM"


def load_iv(cond_folder, dur_dir, fname):
    if cond_folder in OVERRIDE_DIRS:
        folder = os.path.join(OVERRIDE_DIRS[cond_folder], dur_dir)
    else:
        base = VC_POLYOL_OVERRIDE if "VC Polyol" in fname \
               else os.path.join(BASE, "Overpotnital")
        folder = os.path.join(base, cond_folder, dur_dir, "Edited")
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
            if not pts:
                return None, None
            I_d = np.array([a * 1000.0 / AREA for a, _ in pts])  # mA/cm^2
            V   = np.array([v for _, v in pts])
            return I_d, V
    return None, None


# ── Shared y (power) and x (current) maxima across the 4 conditions ──────────
pd_max, i_max = 0.0, 0.0
for cond_folder, k75 in ROWS:
    fname = FNAMES[(cond_folder, SAMPLE)]
    for dur, dur_dir in (("BOL", "BOL"), ("75K", k75)):
        I, V = load_iv(cond_folder, dur_dir, fname)
        if I is not None:
            pd_max = max(pd_max, ((I / 1000) * V).max())
            i_max  = max(i_max, (I / 1000).max())
pd_ymax = np.ceil(pd_max / 0.2) * 0.2
x_max   = np.ceil(i_max / 0.5) * 0.5

# ── Build 4x1 figure ────────────────────────────────────────────────────────
fig = plt.figure(figsize=(5.5, 4.0 * len(ROWS)), dpi=300, constrained_layout=True)
gs  = GridSpec(len(ROWS), 1, figure=fig)

for row_idx, (cond_folder, k75) in enumerate(ROWS):
    ax    = fig.add_subplot(gs[row_idx, 0])
    fname = FNAMES[(cond_folder, SAMPLE)]
    dur_dirs = {"BOL": "BOL", "75K": k75}

    for dur in DUR_ORDER:
        I, V = load_iv(cond_folder, dur_dirs[dur], fname)
        if I is None:
            continue
        P = (I / 1000) * V
        ax.plot(I / 1000, P, ls=DUR_META[dur]["ls"],
                color=DUR_META[dur]["color"], lw=1.6, zorder=3)
        idx_mp = int(np.argmax(P))
        ax.plot(I[idx_mp] / 1000, P[idx_mp], marker=DUR_META[dur]["marker"],
                ms=9, zorder=7, linestyle="none", color=DUR_META[dur]["color"],
                markeredgecolor="white", markeredgewidth=1.2)

    ax.set_xlim(0, x_max)
    ax.set_ylim(0, pd_ymax)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=4, steps=[1, 2, 5, 10]))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda y, p: f"{y:.1f}"))
    ax.xaxis.set_major_locator(MultipleLocator(0.5))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, p: f"{x:.1f}"))
    ax.tick_params(direction="out", top=False, right=False, width=1.2, length=5)

    if row_idx != len(ROWS) - 1:
        ax.tick_params(labelbottom=False)

out = os.path.join(OUT_DIR, "IV_PD_KB_BM_AllCond_PD_linear.png")
fig.savefig(out, dpi=300, bbox_inches="tight")
plt.close(fig)
print(f"Saved -> {out}")
