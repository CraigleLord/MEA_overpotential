"""
Alternative Tafel plots for O2 1.5 bar BP.
Regression window: ln|j| = -3.5 to -2.0  (wider than default -3.5 to -2.5).
No per-sample overrides.
Output: Durability I-V figure/tafel plot alt/
         Durability I-V figure/tafel plot alt/Kinetic slope and intercept alt.xlsx
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
BASE  = os.path.dirname(os.path.abspath(__file__))
OUT   = os.path.join(BASE, "Durability I-V figure", "tafel plot alt")
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "font.family":       "Arial",
    "font.size":         9,
    "axes.linewidth":    0.8,
    "xtick.major.width": 0.8, "ytick.major.width": 0.8,
    "xtick.direction":   "in", "ytick.direction": "in",
    "xtick.top":  True,  "ytick.right": True,
    "xtick.major.size":  4,   "ytick.major.size": 4,
    "xtick.minor.size":  2.5, "ytick.minor.size": 2.5,
})

SAMPLES = [
    ("CN BM",     "CN BM 1.5bp O2"),
    ("CN Polyol", "CN Polyol 1.5bp O2"),
    ("KB BM",     "KB BM 1.5bp O2"),
    ("KB Polyol", "KB Polyol 1.5bp O2"),
    ("VC BM",     "VC BM 1.5bp O2"),
    ("VC Polyol", "VC Polyol 1.5bp O2"),
    ("AB BM",     "AB BM 1.5bp O2"),
    ("AB Polyol", "AB Polyol 1.5bp O2"),
]

# BOL/30K/75K: marker color, regression line color, line style, folder, marker
DURS = [
    ("BOL", "BOL",  "#000000", "red",   "-",  "s"),
    ("30K", "30K",  "#555555", "blue",  "--", "o"),
    ("75K", "75k",  "#999999", "green", ":",  "^"),   # 75k lowercase for O2 BP
]

LN_LO, LN_HI = -3.5, -2.0   # wider regression window

# Per-sample-durability overrides
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


print(f"{'Sample':<12} {'Dur':<4} {'Points':>6} {'Slope':>10} {'Intercept':>10}")
print("-" * 50)

# Collect results for Excel export
results = []   # (samp_lbl, dur_lbl, slope, intercept)

for samp_lbl, fname in SAMPLES:
    fig, ax = plt.subplots(figsize=(4.5, 4.0), dpi=300)

    ann_lines = []
    all_eta   = []

    for dur_lbl, dur_dir, mcolor, lcolor, lst, mkr in DURS:
        j, V = load_iv(dur_dir, fname)
        if j is None:
            print(f"{samp_lbl:<12} {dur_lbl:<4} MISSING")
            results.append((samp_lbl, dur_lbl, None, None))
            continue

        eta  = V - E_EQ
        slope, intercept, ln_j = tafel_fit(j, eta, samp_lbl, dur_lbl)
        all_eta.extend(eta.tolist())

        lo, hi = WINDOW_OVERRIDE.get((samp_lbl, dur_lbl), (LN_LO, LN_HI))
        n_pts = int(((np.log(j) >= lo) & (np.log(j) <= hi)).sum())

        # data points
        ax.plot(ln_j, eta, marker=mkr, color=mcolor, markersize=2.5,
                linestyle="none", markeredgewidth=0, zorder=3, label=dur_lbl)

        # regression line
        if slope is not None:
            x_line = np.linspace(-5, 1, 300)
            ax.plot(x_line, slope * x_line + intercept,
                    color=lcolor, ls="--", lw=1.3, zorder=4)
            ann_lines.append(
                f"{dur_lbl}:  slope={slope:.4f},  intercept={intercept:.3f}"
            )
            print(f"{samp_lbl:<12} {dur_lbl:<4} {n_pts:>6} {slope:>10.4f} {intercept:>10.3f}")
            results.append((samp_lbl, dur_lbl, slope, intercept))
        else:
            print(f"{samp_lbl:<12} {dur_lbl:<4}  TOO FEW POINTS IN TAFEL WINDOW")
            results.append((samp_lbl, dur_lbl, None, None))

    # axes
    y_min = min(all_eta) if all_eta else -1.05
    y_min = np.floor(y_min * 10) / 10 - 0.05
    ax.set_xlim(-5, 1)
    ax.set_ylim(y_min, -0.30)
    ax.set_xlabel(r"ln|j| (j in A/cm$^2$)", fontsize=10)
    ax.set_ylabel(r"Total Overpotential ($\eta$)", fontsize=10)
    ax.set_title(f"{samp_lbl} — O$_2$, 1.5 bar$_g$", fontsize=9, pad=6)
    ax.xaxis.set_major_locator(plt.MultipleLocator(1))
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    ax.tick_params(which="both", direction="in", top=True, right=True)
    ax.minorticks_on()

    ax.legend(loc="lower left", fontsize=8, markerscale=1.6,
              handlelength=0.8, borderpad=0.5, labelspacing=0.3)

    ax.text(0.97, 0.97, "\n".join(ann_lines),
            transform=ax.transAxes, ha="right", va="top",
            fontsize=7, linespacing=1.7,
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.85))

    plt.tight_layout()
    safe = samp_lbl.replace(" ", "")
    out  = os.path.join(OUT, f"Tafel_{safe}_O2BP_alt.png")
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"  -> Saved: {os.path.basename(out)}\n")


# ── Export kinetics Excel ─────────────────────────────────────────────────────
wb_out = openpyxl.Workbook()
ws_out = wb_out.active
ws_out.title = "Sheet1"
ws_out.append([None, None, "Slope", "Intercept"])

prev_samp = None
for samp_lbl, dur_lbl, slope, intercept in results:
    samp_cell = samp_lbl if samp_lbl != prev_samp else None
    prev_samp = samp_lbl
    ws_out.append([
        samp_cell,
        dur_lbl,
        round(slope, 4)     if slope     is not None else None,
        round(intercept, 3) if intercept is not None else None,
    ])

excel_out = os.path.join(OUT, "Kinetic slope and intercept alt.xlsx")
wb_out.save(excel_out)
wb_out.close()
print(f"Excel saved: {excel_out}")
