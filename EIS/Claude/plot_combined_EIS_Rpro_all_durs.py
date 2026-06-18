"""
Combined EIS Nyquist + R_pro bar chart — BOL, 30K, and 75K.
Layout: 2 rows × 2 cols  (Main left, Other right)
  Row 0: EIS Nyquist with all 3 durations + dotted regression lines
  Row 1: Grouped bar chart (BOL / 30K / 75K) of computed R_pro
Output: 260603 New data set (reproducibility)/Final Plots/Combined_EIS_Rpro_AllDur.png
"""

import os, sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib.lines as mlines
import matplotlib.patches as mpatches

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE   = os.path.dirname(os.path.abspath(__file__))
XLFILE = os.path.join(BASE, "Raw data EIS+HFR.xlsx")

ALL_SAMPLES = ["CN BM", "KB BM", "VC Polyol", "AB Polyol",
               "CN Polyol", "KB Polyol", "VC BM", "AB BM"]
GROUPS = {
    "Main":  ["CN BM", "KB BM", "VC Polyol", "AB Polyol"],
    "Other": ["CN Polyol", "KB Polyol", "VC BM", "AB BM"],
}
LABELS = {
    "CN BM": "CN-BM", "KB BM": "KB-BM",
    "VC Polyol": "VC-Polyol", "AB Polyol": "AB-Polyol",
    "CN Polyol": "CN-Polyol", "KB Polyol": "KB-Polyol",
    "VC BM": "VC-BM", "AB BM": "AB-BM",
}
SAMPLE_COLORS = {
    "CN BM":     "#2e7d32", "KB BM":     "#87ceeb",
    "VC Polyol": "#000000", "AB Polyol": "#e6a817",
    "CN Polyol": "#2e7d32", "KB Polyol": "#87ceeb",
    "VC BM":     "#000000", "AB BM":     "#e6a817",
}

DURS = ["BOL", "30K", "75K"]

# Data curve styles: linestyle encodes duration
DUR_LINE = {
    "BOL": dict(ls="-",   lw=1.8, marker="o", ms=5),
    "30K": dict(ls="-.",  lw=1.8, marker="s", ms=5),
    "75K": dict(ls="--",  lw=1.8, marker="^", ms=5),
}

# Bar colours: light → dark encodes degradation
DUR_BAR_COLOR = {
    "BOL": "#cccccc",
    "30K": "#888888",
    "75K": "#444444",
}

FS_LABEL  = 26
FS_TICK   = 22
FS_LEGEND = 18

_AX_XLIM = (0.030, 0.050)
_AX_YLIM = (0.000, 0.080)

# Per-sample regression-window overrides (same as original script)
_REG_N = {
    "CN BM":      4,  "CN Polyol":  4,
    "KB BM":     12,  "KB Polyol": 12,
    "AB Polyol": 10,  "AB BM":     10,
}
_REG_XRANGE = {
    "VC Polyol": (0.043, 0.047),
    "VC BM":     (0.043, 0.047),
}
_REG_YRANGE = {
    "VC Polyol": (-0.05, 0.10),
    "VC BM":     (-0.05, 0.10),
}


# ── Data loading ──────────────────────────────────────────────────────────────
def load_eis_all():
    raw = pd.read_excel(XLFILE, sheet_name="EIS", header=None)
    data = {}
    for si, samp in enumerate(ALL_SAMPLES):
        base = si * 7
        data[samp] = {}
        for di, dur in enumerate(DURS):          # BOL=0, 30K=1, 75K=2
            col = base + di * 2
            zr = raw.iloc[2:, col].astype(float).values
            zi = raw.iloc[2:, col + 1].astype(float).values
            mask = np.isfinite(zr) & np.isfinite(zi)
            data[samp][dur] = (zr[mask], zi[mask])
    return data


def _best_linear_region(zr, zi, win=15):
    n, win = len(zr), min(15, len(zr))
    best_r2, best_c = -np.inf, None
    for i in range(n - win + 1):
        xw, yw = zr[i:i+win], zi[i:i+win]
        c = np.polyfit(xw, yw, 1)
        ss_res = np.sum((yw - np.polyval(c, xw))**2)
        ss_tot = np.sum((yw - yw.mean())**2)
        r2 = 1 - ss_res/ss_tot if ss_tot > 1e-15 else 0
        if r2 > best_r2:
            best_r2, best_c = r2, c
    return best_c


def _reg_coeffs(zr, zi, samp):
    if samp in _REG_XRANGE:
        lo, hi = _REG_XRANGE[samp]
        mask = (zr >= lo) & (zr <= hi) & (zi >= 0)
        if mask.sum() >= 2:
            return np.polyfit(zr[mask], zi[mask], 1)
    if samp in _REG_N:
        vis = ((zr >= _AX_XLIM[0]) & (zr <= _AX_XLIM[1]) &
               (zi >= _AX_YLIM[0]) & (zi <= _AX_YLIM[1]))
        zr_v, zi_v = zr[vis], zi[vis]
        idx = np.argsort(zi_v)
        zr_v, zi_v = zr_v[idx], zi_v[idx]
        n = min(_REG_N[samp], len(zr_v))
        return np.polyfit(zr_v[-n:], zi_v[-n:], 1)
    return _best_linear_region(zr, zi)


def calc_rpro(zr, zi, samp):
    c = _reg_coeffs(zr, zi, samp)
    x_reg   = -c[1] / c[0]
    pos     = np.where(zi > 0)[0]
    x_first = zr[pos[0]] if len(pos) else zr[0]
    return 3.0 * (x_reg - x_first)


# ── Compute everything ────────────────────────────────────────────────────────
eis  = load_eis_all()
rpro = {samp: {dur: calc_rpro(*eis[samp][dur], samp)
               for dur in DURS}
        for samp in ALL_SAMPLES}

# ── Figure ────────────────────────────────────────────────────────────────────
fig, axes = plt.subplots(2, 2, figsize=(18, 14))

# ── Row 0: Nyquist ───────────────────────────────────────────────────────────
for col, (gname, samples) in enumerate(GROUPS.items()):
    ax = axes[0, col]
    for samp in samples:
        color = SAMPLE_COLORS[samp]
        for dur in DURS:
            zr, zi = eis[samp][dur]
            ax.plot(zr, zi, color=color, **DUR_LINE[dur],
                    markerfacecolor=color, markeredgewidth=0.3)
            # Dotted regression line (same color, clearly a fit)
            c = _reg_coeffs(zr, zi, samp)
            x0 = min(zr.min(), -c[1]/c[0], (0.08 - c[1])/c[0])
            x1 = max(zr.max(), -c[1]/c[0], (0.08 - c[1])/c[0])
            xf = np.linspace(x0, x1, 300)
            yf = np.polyval(c, xf)
            ylo, yhi = _REG_YRANGE.get(samp, (0, 0.080))
            vis = (yf >= ylo) & (yf <= yhi)
            ax.plot(xf[vis], yf[vis], color=color, lw=1.6, ls=":", zorder=5)

    ax.set_xlim(*_AX_XLIM)
    ax.set_ylim(*_AX_YLIM)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(0.01))
    ax.xaxis.set_minor_locator(ticker.MultipleLocator(0.005))
    ax.yaxis.set_major_locator(ticker.MultipleLocator(0.02))
    ax.yaxis.set_minor_locator(ticker.MultipleLocator(0.01))
    ax.tick_params(which="major", labelsize=FS_TICK, width=1.5, length=6)
    ax.tick_params(which="minor", width=1.0, length=3)
    ax.set_xlabel("Z' (Ω)", fontsize=FS_LABEL)

    # Sample legend (colored lines)
    samp_handles = [mlines.Line2D([], [], color=SAMPLE_COLORS[s], lw=2,
                                  label=LABELS[s]) for s in samples]
    leg_samp = ax.legend(handles=samp_handles, fontsize=FS_LEGEND,
                         frameon=False, handlelength=1.4, labelspacing=0.3,
                         loc="upper left")
    ax.add_artist(leg_samp)

    # Duration legend (gray lines, linestyle only) — top-right of each panel
    dur_handles = [mlines.Line2D([], [], color="#555555", lw=2,
                                 ls=DUR_LINE[d]["ls"], marker=DUR_LINE[d]["marker"],
                                 ms=7, label=d) for d in DURS]
    ax.legend(handles=dur_handles, fontsize=FS_LEGEND - 2,
              frameon=True, framealpha=0.85, edgecolor="#cccccc",
              handlelength=2.0, labelspacing=0.3, loc="upper right")

axes[0, 0].set_ylabel("Z'' (Ω)", fontsize=FS_LABEL)

# ── Row 1: R_pro grouped bar chart ───────────────────────────────────────────
bar_w   = 0.22
offsets = np.array([-bar_w, 0.0, bar_w])
x_gap   = 1.6   # centre-to-centre spacing between sample groups

all_rpro_vals = [rpro[s][d] for s in ALL_SAMPLES for d in DURS]
bar_ymax = max(all_rpro_vals) * 1.30

for col, (gname, samples) in enumerate(GROUPS.items()):
    ax = axes[1, col]
    n  = len(samples)
    x_centers = np.arange(n) * x_gap

    for di, dur in enumerate(DURS):
        vals = [rpro[s][dur] for s in samples]
        ax.bar(x_centers + offsets[di], vals, width=bar_w,
               facecolor=DUR_BAR_COLOR[dur], edgecolor="black",
               linewidth=0.6, label=dur)

    ax.set_xticks(x_centers)
    ax.set_xticklabels([LABELS[s] for s in samples], fontsize=FS_TICK)
    ax.set_xlim(-0.6, x_centers[-1] + 0.6)
    ax.set_ylim(0, bar_ymax)
    ax.yaxis.set_major_locator(ticker.MultipleLocator(0.004))
    ax.yaxis.set_minor_locator(ticker.MultipleLocator(0.002))
    ax.tick_params(which="major", labelsize=FS_TICK, width=1.5, length=6)
    ax.tick_params(which="minor", width=1.0, length=3)

    for lbl, samp in zip(ax.get_xticklabels(), samples):
        lbl.set_color(SAMPLE_COLORS[samp])
        lbl.set_fontweight("bold")

    if col == 0:
        ax.legend(fontsize=FS_LEGEND, frameon=False,
                  handlelength=1.2, labelspacing=0.3)

axes[1, 0].set_ylabel("Proton Overpotential (V)", fontsize=FS_LABEL)

plt.tight_layout(h_pad=4.0, w_pad=3.0)

OUT_DIR = os.path.join(BASE, "..", "..", "260603 New data set (reproducibility)",
                       "Final Plots")
os.makedirs(OUT_DIR, exist_ok=True)
out = os.path.join(OUT_DIR, "Combined_EIS_Rpro_AllDur.png")
fig.savefig(out, dpi=200, bbox_inches="tight")
plt.close(fig)
print(f"Saved -> {out}")
