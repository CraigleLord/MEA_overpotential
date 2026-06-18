"""
Individual (un-gridded) I-V / Power-density / waterfall-bar subplots — one PNG
per panel, with all axis titles removed. Re-renders the panels that make up
LSV_combined_2x4_{Main,Other}_{0p4V,0p7V}.png (plot_lsv_combined_2x4.py),
splitting each condition row into its own self-contained figure (per-row
condition annotation, durability/sample legends, and bar x-tick labels
included on every panel).

Output ->
  260603 New data set (reproducibility)/Final Plots/LSV_combined_2x4_Main_0p4V/
  260603 New data set (reproducibility)/Final Plots/LSV_combined_2x4_Main_0p7V/
  260603 New data set (reproducibility)/Final Plots/LSV_combined_2x4_Other_0p4V/
  260603 New data set (reproducibility)/Final Plots/LSV_combined_2x4_Other_0p7V/
"""

import os, sys, openpyxl
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.lines as mlines
from matplotlib.patches import Patch
from matplotlib.gridspec import GridSpec

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE      = os.path.dirname(os.path.abspath(__file__))
AREA      = 5.0
REPRO_DIR = os.path.join(BASE, "260603 New data set (reproducibility)")
FINAL_DIR = os.path.join(REPRO_DIR, "Final Plots")

FS_TICK   = 22
FS_LABEL  = 22
FS_LEGEND = 22
FS_ANNOT  = 22

plt.rcParams.update({
    "font.family":       "Arial",
    "font.size":         FS_TICK,
    "axes.linewidth":    1.0,
    "xtick.major.width": 1.2, "ytick.major.width": 1.2,
    "xtick.minor.width": 0.8, "ytick.minor.width": 0.8,
    "xtick.direction":   "out", "ytick.direction": "out",
    "xtick.top":  False, "ytick.right": False,
    "xtick.major.size":  5,   "ytick.major.size":  5,
    "xtick.minor.size":  3,   "ytick.minor.size":  3,
    "legend.frameon":    True, "legend.framealpha": 0.9,
    "legend.edgecolor":  "none",
    "lines.linewidth":   1.6,
})

# ── Sample groups (matching plot_lsv_combined_2x4.py) ──────────────────────
MAIN_CLR = {
    "CN BM":     "#1f7a1f",
    "KB BM":     "#56b4e9",
    "VC Polyol": "#000000",
    "AB Polyol": "#d4a017",
}
OTHER_CLR = {
    "CN Polyol": "#1f7a1f",
    "KB Polyol": "#56b4e9",
    "VC BM":     "#000000",
    "AB BM":     "#e69f00",
}
MAIN_ORDER  = ["CN BM",     "KB BM",     "VC Polyol", "AB Polyol"]
OTHER_ORDER = ["CN Polyol", "KB Polyol", "VC BM",     "AB BM"]

# ── File-name lookup (matching plot_lsv_combined_2x4.py) ───────────────────
FNAMES = {
    ("O2",    "CN BM"):     "CN BM o2",
    ("O2",    "CN Polyol"): "CN Polyol O2",
    ("O2",    "KB BM"):     "KB BM O2",
    ("O2",    "KB Polyol"): "KB Polyol O2",
    ("O2",    "VC BM"):     "VC BM O2",
    ("O2",    "VC Polyol"): "VC Polyol O2",
    ("O2",    "AB BM"):     "AB BM  O2",
    ("O2",    "AB Polyol"): "AB Polyol O2",

    ("O2 BP", "CN BM"):     "CN BM 1.5bp O2",
    ("O2 BP", "CN Polyol"): "CN Polyol 1.5bp O2",
    ("O2 BP", "KB BM"):     "KB BM 1.5bp O2",
    ("O2 BP", "KB Polyol"): "KB Polyol 1.5bp O2",
    ("O2 BP", "VC BM"):     "VC BM 1.5bp O2",
    ("O2 BP", "VC Polyol"): "VC Polyol 1.5bp O2",
    ("O2 BP", "AB BM"):     "AB BM 1.5bp O2",
    ("O2 BP", "AB Polyol"): "AB Polyol 1.5bp O2",

    ("Air",   "CN BM"):     "CN BM air",
    ("Air",   "CN Polyol"): "CN Polyol air",
    ("Air",   "KB BM"):     "KB BM air",
    ("Air",   "KB Polyol"): "KB Polyol air",
    ("Air",   "VC BM"):     "VC BM air",
    ("Air",   "VC Polyol"): "VC Polyol air",
    ("Air",   "AB BM"):     "AB BM Air",
    ("Air",   "AB Polyol"): "AB Polyol air",

    ("Air BP","CN BM"):     "CN BM 1.5bp air",
    ("Air BP","CN Polyol"): "CN Polyol 1.5bp air",
    ("Air BP","KB BM"):     "KB BM 1.5bp air",
    ("Air BP","KB Polyol"): "KB Polyol 1.5bp air",
    ("Air BP","VC BM"):     "VC BM 1.5bp air",
    ("Air BP","VC Polyol"): "VC Polyol 1.5bp air",
    ("Air BP","AB BM"):     "AB BM 1.5bp air",
    ("Air BP","AB Polyol"): "AB Polyol 1.5bp air",
}

ROWS = [
    ("Air",    "H$_2$/Air",   "0 bar$_g$",   "75K"),
    ("Air BP", "H$_2$/Air",   "1.5 bar$_g$", "75K"),
    ("O2",     "H$_2$/O$_2$", "0 bar$_g$",   "75K"),
    ("O2 BP",  "H$_2$/O$_2$", "1.5 bar$_g$", "75k"),
]
COND_TAG = {
    "Air":    "Air_0bp",
    "Air BP": "Air_15bp",
    "O2":     "O2_0bp",
    "O2 BP":  "O2_15bp",
}

XLIM_ALL = 3500
PMAX_ALL = 1000

BAR_DUR_ORDER = ["75K", "30K", "BOL"]
BAR_STYLES = {
    "BOL": dict(facecolor="#cccccc", hatch="",    edgecolor="black", linewidth=0.5),
    "30K": dict(facecolor="#999999", hatch="///", edgecolor="black", linewidth=0.5),
    "75K": dict(facecolor="#555555", hatch="",    edgecolor="black", linewidth=0.5),
}
BAR_LABELS = {"BOL": "0 K", "30K": "30 K", "75K": "75 K"}

PANEL_FIGSIZE_IV  = (5.5, 4.5)
PANEL_FIGSIZE_BAR = (5.0, 4.8)


def load_iv(cond_folder, dur_dir, fname):
    folder = os.path.join(BASE, "Overpotnital", cond_folder, dur_dir, "Edited")
    for sfx in ("_edited.xlsx", "_unchanged.xlsx"):
        p = os.path.join(folder, fname + sfx)
        if os.path.exists(p):
            wb  = openpyxl.load_workbook(p, data_only=True)
            ws  = wb.active
            pts = [(ws.cell(r, 1).value, ws.cell(r, 2).value)
                   for r in range(4, ws.max_row + 1)
                   if isinstance(ws.cell(r, 1).value, (int, float))
                   and isinstance(ws.cell(r, 2).value, (int, float))]
            wb.close()
            if not pts:
                return None, None, None
            I_d = np.array([a * 1000.0 / AREA for a, _ in pts])
            V   = np.array([v for _, v in pts])
            return I_d, V, I_d * V
    return None, None, None


def get_i_at_v(I, V, v_target):
    if I is None or V is None or len(I) < 2:
        return np.nan
    order = np.argsort(V)
    Vs, Is = V[order], I[order]
    if v_target < Vs[0] or v_target > Vs[-1]:
        return np.nan
    return float(np.interp(v_target, Vs, Is))


# ── Shared bar y-scale, computed from MAIN_ORDER (matches original) ────────
def compute_row_bar_increments(cond_folder, k75, sample_order, v_target):
    dur_dirs = {"BOL": "BOL", "30K": "30K", "75K": k75}
    i_abs = {}
    for dur_key in ["BOL", "30K", "75K"]:
        vals = []
        for samp in sample_order:
            fname = FNAMES[(cond_folder, samp)]
            I, V, _ = load_iv(cond_folder, dur_dirs[dur_key], fname)
            v = get_i_at_v(I, V, v_target)
            vals.append(0.0 if np.isnan(v) else v)
        i_abs[dur_key] = np.array(vals)
    increments = {
        "75K": i_abs["75K"],
        "30K": np.maximum(i_abs["30K"] - i_abs["75K"], 0),
        "BOL": np.maximum(i_abs["BOL"] - i_abs["30K"], 0),
    }
    return increments, i_abs["BOL"].max()


def compute_global_bar_max(v_target):
    totals = []
    for cond_folder, _, _, k75 in ROWS:
        _, tot_main  = compute_row_bar_increments(cond_folder, k75, MAIN_ORDER, v_target)
        _, tot_other = compute_row_bar_increments(cond_folder, k75, OTHER_ORDER, v_target)
        totals += [tot_main, tot_other]
    return max(totals)


# Extra headroom so the BOL/30K/75K legend on every panel doesn't overlap
# the tallest waterfall bar (legend occupies ~30% of the panel height).
LEGEND_HEADROOM = 0.40
YMAX = {v: np.ceil((compute_global_bar_max(v) / (1 - LEGEND_HEADROOM)) / 100) * 100
        for v in (0.70, 0.40)}


# ── Per-panel drawing (axis titles intentionally omitted) ──────────────────
def draw_iv(ax, cond_folder, gas_str, bp_str, k75, sample_order, color_map):
    durs = [("BOL", "BOL", "-"), ("30K", "30K", "--"), ("75K", k75, ":")]
    for samp in sample_order:
        clr   = color_map[samp]
        fname = FNAMES[(cond_folder, samp)]
        for _, dur_dir, ls in durs:
            I, V, _ = load_iv(cond_folder, dur_dir, fname)
            if I is None:
                continue
            ax.plot(I, V, ls=ls, color=clr, lw=1.6)

    ax.set_xlim(0, XLIM_ALL)
    ax.minorticks_on()
    ax.xaxis.set_major_locator(plt.MultipleLocator(1000))
    ax.xaxis.set_minor_locator(plt.MultipleLocator(500))
    ax.set_ylim(0.2, 1.0)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.2))

    annot = (
        f"{gas_str}\n"
        r"0.05 mg$_\mathregular{Pt}$/cm$^2$"
        f"\n{bp_str}"
    )
    ax.text(0.97, 0.97, annot,
            transform=ax.transAxes, ha="right", va="top",
            fontsize=FS_ANNOT, linespacing=1.45)

    dur_handles = [
        mlines.Line2D([], [], color="gray", ls="-",  lw=1.6, label="BOL"),
        mlines.Line2D([], [], color="gray", ls="--", lw=1.6, label="30K"),
        mlines.Line2D([], [], color="gray", ls=":",  lw=1.6, label="75K"),
    ]
    ax.legend(handles=dur_handles, loc="lower right",
              fontsize=FS_LEGEND, handlelength=2.5,
              borderpad=0.5, labelspacing=0.3)


def draw_pd(ax, cond_folder, k75, sample_order, color_map):
    durs = [("BOL", "BOL", "-"), ("30K", "30K", "--"), ("75K", k75, ":")]
    for samp in sample_order:
        clr   = color_map[samp]
        fname = FNAMES[(cond_folder, samp)]
        for _, dur_dir, ls in durs:
            I, V, P = load_iv(cond_folder, dur_dir, fname)
            if I is None:
                continue
            ax.plot(I, P, ls=ls, color=clr, lw=1.6)

    ax.set_xlim(0, XLIM_ALL)
    ax.minorticks_on()
    ax.xaxis.set_major_locator(plt.MultipleLocator(1000))
    ax.xaxis.set_minor_locator(plt.MultipleLocator(500))
    ax.set_ylim(0, PMAX_ALL)
    ax.yaxis.set_major_locator(plt.MultipleLocator(200))

    samp_handles = [
        mlines.Line2D([], [], color=color_map[s], ls="-", lw=1.8, label=s)
        for s in sample_order
    ]
    ax.legend(handles=samp_handles, loc="lower right",
              fontsize=FS_LEGEND, handlelength=1.6,
              borderpad=0.5, labelspacing=0.3)


def draw_bar(ax, cond_folder, k75, sample_order, v_target):
    increments, _ = compute_row_bar_increments(cond_folder, k75, sample_order, v_target)
    x_pos = np.arange(len(sample_order))

    bottoms = np.zeros(len(sample_order))
    for dur_key in BAR_DUR_ORDER:
        ax.bar(x_pos, increments[dur_key], bottom=bottoms, width=0.55,
               label=BAR_LABELS[dur_key], **BAR_STYLES[dur_key])
        bottoms += increments[dur_key]

    ax.set_xticks(x_pos)
    ax.set_xticklabels(sample_order, fontsize=FS_TICK, rotation=40, ha="right")
    ax.tick_params(axis="y", labelsize=FS_TICK)
    ax.tick_params(axis="x", which="minor", bottom=False)
    ax.minorticks_on()
    ax.set_ylim(0, YMAX[v_target])

    leg_handles = [Patch(label=BAR_LABELS[d], **BAR_STYLES[d]) for d in ["BOL", "30K", "75K"]]
    ax.legend(handles=leg_handles, loc="upper right",
              fontsize=FS_LEGEND, frameon=False,
              handlelength=1.2, labelspacing=0.3)


# ── Figure-per-panel builders ───────────────────────────────────────────────
def make_panel(figsize, draw_fn, *args):
    fig = plt.figure(figsize=figsize, dpi=300, constrained_layout=True)
    gs = GridSpec(1, 1, figure=fig)
    ax = fig.add_subplot(gs[0, 0])
    draw_fn(ax, *args)
    return fig


def save_panel(fig, out_dir, name):
    out = os.path.join(out_dir, name)
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {name}")


def split_combined(sample_order, color_map, group_name, v_target, suffix):
    out_dir = os.path.join(FINAL_DIR, f"LSV_combined_2x4_{group_name}{suffix}")
    os.makedirs(out_dir, exist_ok=True)

    for cond_folder, gas_str, bp_str, k75 in ROWS:
        tag = COND_TAG[cond_folder]

        fig = make_panel(PANEL_FIGSIZE_IV, draw_iv,
                          cond_folder, gas_str, bp_str, k75, sample_order, color_map)
        save_panel(fig, out_dir, f"LSV_combined_2x4_{group_name}{suffix}_{tag}_IV.png")

        fig = make_panel(PANEL_FIGSIZE_IV, draw_pd,
                          cond_folder, k75, sample_order, color_map)
        save_panel(fig, out_dir, f"LSV_combined_2x4_{group_name}{suffix}_{tag}_PD.png")

        fig = make_panel(PANEL_FIGSIZE_BAR, draw_bar,
                          cond_folder, k75, sample_order, v_target)
        save_panel(fig, out_dir, f"LSV_combined_2x4_{group_name}{suffix}_{tag}_Bar.png")


split_combined(MAIN_ORDER,  MAIN_CLR,  "Main",  0.70, "_0p7V")
split_combined(MAIN_ORDER,  MAIN_CLR,  "Main",  0.40, "_0p4V")
split_combined(OTHER_ORDER, OTHER_CLR, "Other", 0.70, "_0p7V")
split_combined(OTHER_ORDER, OTHER_CLR, "Other", 0.40, "_0p4V")

print("Done.")
