"""
Individual (un-gridded) waterfall bar-chart subplots — one PNG per panel,
with the shared y-axis title removed. Re-renders the panels that make up:
  Bar_Other_0p7_0p4V_grid.png  (build_bar_only_grid, OTHER_ORDER, 2 cols x 4 rows)
from plot_lsv_kb_dur_grid.py, splitting each row x column into its own
self-contained figure: sample-name x-tick labels, "j @ V" title, BOL/75K
legend, and condition (gas/pressure) text are included on every panel.

Output -> 260603 New data set (reproducibility)/Final Plots/Other 0.7_0.4V bar chart/
"""

import os, sys, openpyxl
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.gridspec import GridSpec
from matplotlib.ticker import MaxNLocator, FuncFormatter, MultipleLocator

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE          = os.path.dirname(os.path.abspath(__file__))
AREA          = 5.0
REPRO_DIR     = os.path.join(BASE, "260603 New data set (reproducibility)")
OUT_DIR       = os.path.join(REPRO_DIR, "Final Plots", "Other 0.7_0.4V bar chart")
MAIN_OUT_DIR  = os.path.join(REPRO_DIR, "Final Plots", "Main_barchart")
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(MAIN_OUT_DIR, exist_ok=True)

# ── 260603 staging override for VC Polyol ──────────────────────────────────
VC_POLYOL_OVERRIDE = os.path.join(REPRO_DIR, "Overpotnital")

# ── Unified font size — everything at the same size ────────────────────────
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
    "xtick.minor.size":  3,   "ytick.minor.size":  3,
    "lines.linewidth":   1.6,
    "legend.fontsize":   FS,
    "axes.labelsize":    FS,
    "xtick.labelsize":   FS,
    "ytick.labelsize":   FS,
})

# ── Conditions (rows) ──────────────────────────────────────────────────────
ROWS = [
    ("Air",   "H$_2$/Air",   "0 bar$_g$",   "75K"),
    ("Air BP","H$_2$/Air",   "1.5 bar$_g$", "75K"),
    ("O2",    "H$_2$/O$_2$", "0 bar$_g$",   "75K"),
    ("O2 BP", "H$_2$/O$_2$", "1.5 bar$_g$", "75k"),
]

COND_TAG = {
    "Air":    "Air_0bp",
    "Air BP": "Air_15bp",
    "O2":     "O2_0bp",
    "O2 BP":  "O2_15bp",
}

# ── Sample groups (matching plot_lsv_kb_dur_grid.py) ───────────────────────
MAIN_ORDER  = ["CN BM",     "KB BM",     "VC Polyol", "AB Polyol"]
OTHER_ORDER = ["CN Polyol", "KB Polyol", "VC BM",     "AB BM"]

# ── File stems (matching plot_lsv_kb_dur_grid.py) ──────────────────────────
FNAMES = {
    ("Air",   "CN BM"):     "CN BM air",
    ("Air",   "KB BM"):     "KB BM air",
    ("Air",   "VC Polyol"): "VC Polyol air",
    ("Air",   "AB Polyol"): "AB Polyol air",
    ("Air",   "CN Polyol"): "CN Polyol air",
    ("Air",   "KB Polyol"): "KB Polyol air",
    ("Air",   "VC BM"):     "VC BM air",
    ("Air",   "AB BM"):     "AB BM Air",

    ("Air BP","CN BM"):     "CN BM 1.5bp air",
    ("Air BP","KB BM"):     "KB BM 1.5bp air",
    ("Air BP","VC Polyol"): "VC Polyol 1.5bp air",
    ("Air BP","AB Polyol"): "AB Polyol 1.5bp air",
    ("Air BP","CN Polyol"): "CN Polyol 1.5bp air",
    ("Air BP","KB Polyol"): "KB Polyol 1.5bp air",
    ("Air BP","VC BM"):     "VC BM 1.5bp air",
    ("Air BP","AB BM"):     "AB BM 1.5bp air",

    ("O2",    "CN BM"):     "CN BM o2",
    ("O2",    "KB BM"):     "KB BM O2",
    ("O2",    "VC Polyol"): "VC Polyol O2",
    ("O2",    "AB Polyol"): "AB Polyol O2",
    ("O2",    "CN Polyol"): "CN Polyol O2",
    ("O2",    "KB Polyol"): "KB Polyol O2",
    ("O2",    "VC BM"):     "VC BM O2",
    ("O2",    "AB BM"):     "AB BM  O2",

    ("O2 BP", "CN BM"):     "CN BM 1.5bp O2",
    ("O2 BP", "KB BM"):     "KB BM 1.5bp O2",
    ("O2 BP", "VC Polyol"): "VC Polyol 1.5bp O2",
    ("O2 BP", "AB Polyol"): "AB Polyol 1.5bp O2",
    ("O2 BP", "CN Polyol"): "CN Polyol 1.5bp O2",
    ("O2 BP", "KB Polyol"): "KB Polyol 1.5bp O2",
    ("O2 BP", "VC BM"):     "VC BM 1.5bp O2",
    ("O2 BP", "AB BM"):     "AB BM 1.5bp O2",
}

# ── Bar chart style (matching plot_lsv_kb_dur_grid.py) ─────────────────────
DUR_ORDER     = ["BOL", "75K"]
BAR_DUR_ORDER = ["75K", "BOL"]
BAR_STYLES = {
    "BOL": dict(facecolor="#cccccc", hatch="", edgecolor="black", linewidth=0.5),
    "75K": dict(facecolor="#555555", hatch="", edgecolor="black", linewidth=0.5),
}
BAR_LABELS = {"BOL": "0 K", "75K": "75 K"}

# ── Voltage reference columns ──────────────────────────────────────────────
V_REF  = [0.70, 0.40]
V_TAG  = {0.70: "0p7V", 0.40: "0p4V"}

PANEL_FIGSIZE = (5.5, 4.0)


# Per-stem data overrides (260603 reproducibility staging: files sit directly in
# {dir}/{dur}/, no "Edited" subfolder). Affects only the listed stems.
REPRO_ALT = os.path.join(REPRO_DIR, "Overpotential DATA raw alt")
DATA_OVERRIDE_DIRS = {
    "KB BM 1.5bp air": os.path.join(REPRO_ALT, "Air BP", "KB BM"),
}


# ── Helpers ────────────────────────────────────────────────────────────────
def load_iv(cond_folder, dur_dir, fname, override_vc=False):
    if fname in DATA_OVERRIDE_DIRS:
        folder = os.path.join(DATA_OVERRIDE_DIRS[fname], dur_dir)
    else:
        base = VC_POLYOL_OVERRIDE if (override_vc and "VC Polyol" in fname) \
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
            I_d = np.array([a * 1000.0 / AREA for a, _ in pts])
            V   = np.array([v for _, v in pts])
            return I_d, V
    return None, None


def j_at_v(I, V, v_target):
    if I is None or len(I) < 2:
        return np.nan
    order = np.argsort(V)
    Vs, Is = V[order], I[order]
    if v_target < Vs[0] or v_target > Vs[-1]:
        return np.nan
    return float(np.interp(v_target, Vs, Is))


# ── Shared bar y-scale across BOTH MAIN/OTHER figures (matches original) ───
def _collect_bar_j(sample_orders):
    bar_j = {v: [] for v in V_REF}
    for sample_order in sample_orders:
        for cond_folder, _, _, k75 in ROWS:
            dur_dirs = {"BOL": "BOL", "30K": "30K", "75K": k75}
            for samp in sample_order:
                fname = FNAMES.get((cond_folder, samp))
                if not fname:
                    continue
                for dur in DUR_ORDER:
                    I, V = load_iv(cond_folder, dur_dirs[dur], fname,
                                   override_vc=True)
                    for v_tgt in V_REF:
                        j = j_at_v(I, V, v_tgt)
                        if not np.isnan(j):
                            bar_j[v_tgt].append(j / 1000)   # A cm⁻²
    steps = {V_REF[0]: 0.2, V_REF[1]: 0.5}
    return {v: np.ceil(max(bar_j[v]) / steps[v]) * steps[v] if bar_j[v] else 2.0
            for v in V_REF}


# Fixed y-axis per voltage column (user-specified):
#   0.7 V → 0–0.8 A cm⁻², 0.2 per tick
#   0.4 V → 0–2.5 A cm⁻², 0.5 per tick
PANEL_YMAXES = {0.70: 0.8, 0.40: 2.5}
PANEL_YSTEP  = {0.70: 0.2, 0.40: 0.5}


# ── Per-panel drawing (y-axis title intentionally omitted) ─────────────────
def draw_bar(ax, cond_folder, gas_str, bp_str, k75, sample_order, v_tgt,
             show_cond_text=True):
    x_pos = np.arange(len(sample_order))

    bol_vals = []
    for samp in sample_order:
        fname = FNAMES.get((cond_folder, samp))
        I, V  = load_iv(cond_folder, "BOL", fname, override_vc=True) if fname else (None, None)
        j = j_at_v(I, V, v_tgt)
        bol_vals.append((0.0 if np.isnan(j) else j) / 1000)   # A cm⁻²

    ax.bar(x_pos, bol_vals, width=0.55, **BAR_STYLES["BOL"])

    # Value label on top of each bar (A cm⁻², 2 d.p.)
    y_off = 0.012 * PANEL_YMAXES[v_tgt]
    for xp, val in zip(x_pos, bol_vals):
        if val > 0:
            ax.text(xp, val + y_off, f"{val:.2f}", ha="center", va="bottom",
                    fontsize=FS - 4)

    ax.set_xlim(-0.5, len(sample_order) - 0.5)
    ax.set_ylim(0, PANEL_YMAXES[v_tgt])
    ax.yaxis.set_major_locator(MultipleLocator(PANEL_YSTEP[v_tgt]))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda y, p: f"{y:.1f}"))
    ax.set_xticks(x_pos)
    ax.set_xticklabels([""] * len(sample_order))
    ax.tick_params(axis="y", labelsize=FS)
    ax.tick_params(axis="x", which="minor", bottom=False)
    ax.minorticks_on()
    ax.tick_params(direction="out", top=False, right=False, width=1.2, length=5)

    if show_cond_text:
        ax.text(0.03, 0.97, f"{gas_str}\n{bp_str}", transform=ax.transAxes,
                ha="left", va="top", fontsize=FS, linespacing=1.45)


# ── Figure-per-panel builders ───────────────────────────────────────────────
def make_panel(*args, **kwargs):
    fig = plt.figure(figsize=PANEL_FIGSIZE, dpi=300, constrained_layout=True)
    gs = GridSpec(1, 1, figure=fig)
    ax = fig.add_subplot(gs[0, 0])
    draw_bar(ax, *args, **kwargs)
    return fig


def save_panel(fig, out_dir, name):
    out = os.path.join(out_dir, name)
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {name}")


for cond_folder, gas_str, bp_str, k75 in ROWS:
    for v_tgt in V_REF:
        show_txt = (v_tgt == 0.70)   # condition text only on 0.7 V panels

        fig = make_panel(cond_folder, gas_str, bp_str, k75, OTHER_ORDER, v_tgt,
                         show_cond_text=show_txt)
        save_panel(fig, OUT_DIR,
                   f"Bar_Other_0p7_0p4V_grid_{COND_TAG[cond_folder]}_{V_TAG[v_tgt]}.png")

        fig = make_panel(cond_folder, gas_str, bp_str, k75, MAIN_ORDER, v_tgt,
                         show_cond_text=show_txt)
        save_panel(fig, MAIN_OUT_DIR,
                   f"Bar_Main_0p7_0p4V_grid_{COND_TAG[cond_folder]}_{V_TAG[v_tgt]}.png")

print("Done.")
