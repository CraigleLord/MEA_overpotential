"""
Individual (un-gridded) EIS Nyquist + R_pro bar-chart subplots — one PNG per
panel, with x/y axis titles removed. Re-renders the 4 panels that make up
Combined_EIS_Rpro.png (EIS/Claude/plot_combined_EIS_Rpro.py), splitting each
Main/Other column into its own self-contained figure (legend / colored
sample x-tick labels are kept on every panel).

Output -> 260603 New data set (reproducibility)/Final Plots/EIS Rpro plots/
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.gridspec import GridSpec

BASE      = os.path.dirname(os.path.abspath(__file__))
EIS_BASE  = os.path.join(BASE, "EIS", "Claude")
XLFILE    = os.path.join(EIS_BASE, "Raw data EIS+HFR.xlsx")
REPRO_DIR = os.path.join(BASE, "260603 New data set (reproducibility)")
OUT_DIR   = os.path.join(REPRO_DIR, "Final Plots", "EIS Rpro plots")
os.makedirs(OUT_DIR, exist_ok=True)

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
# CN=dark green, KB=light blue, VC=black, AB=golden yellow
SAMPLE_COLORS = {
    "CN BM":     "#2e7d32", "KB BM":     "#87ceeb",
    "VC Polyol": "#000000", "AB Polyol": "#e6a817",
    "CN Polyol": "#2e7d32", "KB Polyol": "#87ceeb",
    "VC BM":     "#000000", "AB BM":     "#e6a817",
}

FS_TICK   = 22
FS_LEGEND = 20
BAR_WIDTH = 0.5
PANEL_FIGSIZE = (8, 7)


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


def best_linear_region(zr, zi, win=15):
    """Return (slope, intercept) of the most linear window of `win` points."""
    n = len(zr)
    win = min(win, n)
    best_r2, best_coeffs = -np.inf, None
    for i in range(n - win + 1):
        xw, yw = zr[i:i + win], zi[i:i + win]
        coeffs = np.polyfit(xw, yw, 1)
        y_pred = np.polyval(coeffs, xw)
        ss_res = np.sum((yw - y_pred) ** 2)
        ss_tot = np.sum((yw - yw.mean()) ** 2)
        r2 = 1 - ss_res / ss_tot if ss_tot > 1e-15 else 0
        if r2 > best_r2:
            best_r2 = r2
            best_coeffs = coeffs
    return best_coeffs


eis = load_eis()

# Axis bounds (kept fixed — do not change)
_AX_XLIM = (0.030, 0.050)
_AX_YLIM = (0.000, 0.080)

# Per-sample override: use the last N highest-Z'' points visible within the
# axis bounds as the regression window instead of best_linear_region.
_REG_N = {
    "CN BM":     4,  "CN Polyol": 4,
    "KB BM":    12,  "KB Polyol": 12,
    "AB Polyol": 10, "AB BM":     10,
}

# X-range based regression: use only points with Z' in [x_lo, x_hi]
_REG_XRANGE = {
    "VC Polyol": (0.043, 0.047),
    "VC BM":     (0.043, 0.047),
}

# Per-sample override for how far to draw the regression line (y extents)
_REG_YRANGE = {
    "VC Polyol": (-0.05, 0.10),
    "VC BM":     (-0.05, 0.10),
}


def _reg_coeffs(zr, zi, samp):
    """Return linear fit coefficients for the regression region of `samp`."""
    if samp in _REG_XRANGE:
        x_lo, x_hi = _REG_XRANGE[samp]
        mask = (zr >= x_lo) & (zr <= x_hi) & (zi >= 0)   # exclude inductive branch
        if mask.sum() >= 2:
            return np.polyfit(zr[mask], zi[mask], 1)
    if samp in _REG_N:
        vis = ((zr >= _AX_XLIM[0]) & (zr <= _AX_XLIM[1]) &
               (zi >= _AX_YLIM[0]) & (zi <= _AX_YLIM[1]))
        zr_v, zi_v = zr[vis], zi[vis]
        order = np.argsort(zi_v)
        zr_v, zi_v = zr_v[order], zi_v[order]
        n = min(_REG_N[samp], len(zr_v))
        return np.polyfit(zr_v[-n:], zi_v[-n:], 1)
    return best_linear_region(zr, zi, win=15)


# Compute R_pro from data: 3 x (x_first_arc_point - x_reg_intercept)
_rpro_flat = {}
for _samp, (_zr, _zi) in eis.items():
    _coeffs = _reg_coeffs(_zr, _zi, _samp)
    _x_reg  = -_coeffs[1] / _coeffs[0]
    _pos    = np.where(_zi > 0)[0]
    _x_first = _zr[_pos[0]] if len(_pos) > 0 else _zr[0]
    _rpro_flat[_samp] = 3.0 * (_x_reg - _x_first)

RPRO = {
    "Main":  {s: _rpro_flat[s] for s in GROUPS["Main"]},
    "Other": {s: _rpro_flat[s] for s in GROUPS["Other"]},
}

_BAR_YMAX = max(v for rpro_data in RPRO.values() for v in rpro_data.values()) * 1.25


# ── Per-panel drawing (axis titles intentionally omitted) ──────────────────
def draw_eis(ax, samples):
    for samp in samples:
        zr, zi = eis[samp]
        color = SAMPLE_COLORS[samp]

        ax.plot(zr, zi, color=color, lw=1.5,
                marker="o", markersize=8, markerfacecolor=color,
                label=LABELS[samp])

        coeffs = _reg_coeffs(zr, zi, samp)
        x_at_y0  = -coeffs[1] / coeffs[0]
        x_at_y08 = (0.080 - coeffs[1]) / coeffs[0]
        x_start = min(zr.min(), x_at_y0,  x_at_y08)
        x_end   = max(zr.max(), x_at_y0,  x_at_y08)
        x_fit = np.linspace(x_start, x_end, 300)
        y_fit = np.polyval(coeffs, x_fit)
        y_lo, y_hi = _REG_YRANGE.get(samp, (0, 0.080))
        visible = (y_fit >= y_lo) & (y_fit <= y_hi)
        ax.plot(x_fit[visible], y_fit[visible],
                color=color, lw=2.0, linestyle="--", zorder=5)

    ax.set_xlim(*_AX_XLIM)
    ax.set_ylim(*_AX_YLIM)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(0.01))
    ax.xaxis.set_minor_locator(ticker.MultipleLocator(0.005))
    ax.yaxis.set_major_locator(ticker.MultipleLocator(0.02))
    ax.yaxis.set_minor_locator(ticker.MultipleLocator(0.01))
    ax.tick_params(which="major", labelsize=FS_TICK, width=1.5, length=6)
    ax.tick_params(which="minor", width=1.0, length=3)
    ax.legend(fontsize=FS_LEGEND, frameon=False,
              handlelength=1.4, labelspacing=0.3)


def draw_rpro(ax, rpro_data):
    samples = list(rpro_data.keys())
    x = np.arange(len(samples)) * 1.5   # wider spacing between bars

    vals   = [rpro_data[s] for s in samples]
    colors = [SAMPLE_COLORS[s] for s in samples]
    ax.bar(x, vals, BAR_WIDTH, color=colors, edgecolor="white", linewidth=0.5)

    ax.set_xticks(x)
    ax.set_xticklabels([LABELS[s] for s in samples], fontsize=FS_TICK)
    ax.set_xlim(-0.6, x[-1] + 0.6)
    ax.set_ylim(0, _BAR_YMAX)
    ax.yaxis.set_major_locator(ticker.MultipleLocator(0.004))
    ax.yaxis.set_minor_locator(ticker.MultipleLocator(0.002))
    ax.tick_params(which="major", labelsize=FS_TICK, width=1.5, length=6)
    ax.tick_params(which="minor", width=1.0, length=3)

    for label, samp in zip(ax.get_xticklabels(), samples):
        label.set_color(SAMPLE_COLORS[samp])
        label.set_fontweight("bold")
        label.set_rotation(20)
        label.set_ha("right")


# ── Figure-per-panel builders ───────────────────────────────────────────────
def make_panel(draw_fn, *args):
    fig = plt.figure(figsize=PANEL_FIGSIZE, dpi=200, constrained_layout=True)
    gs = GridSpec(1, 1, figure=fig)
    ax = fig.add_subplot(gs[0, 0])
    draw_fn(ax, *args)
    return fig


def save_panel(fig, name):
    out = os.path.join(OUT_DIR, name)
    fig.savefig(out, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {name}")


for gname, samples in GROUPS.items():
    fig = make_panel(draw_eis, samples)
    save_panel(fig, f"Combined_EIS_Rpro_Nyquist_{gname}.png")

for gname, rpro_data in RPRO.items():
    fig = make_panel(draw_rpro, rpro_data)
    save_panel(fig, f"Combined_EIS_Rpro_Rpro_{gname}.png")

print("Done.")
