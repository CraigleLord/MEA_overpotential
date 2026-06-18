"""
Individual (un-gridded) I-V and Power-Density subplots — one PNG per panel,
with x/y axis labels removed. Re-renders the panels that make up:
  PD_KB_BM_grid.png        (build_pd_grid,    1 col x 4 rows)
  IV_PD_*_grid.png  (x8)   (build_iv_pd_grid, 2 cols x 4 rows)
from plot_lsv_kb_dur_grid.py, splitting each row into its own figure.

Output -> 260603 New data set (reproducibility)/Final Plots/PD plots/
"""

import os, sys, openpyxl
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.lines as mlines
from matplotlib.gridspec import GridSpec
from matplotlib.ticker import MaxNLocator, FuncFormatter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE      = os.path.dirname(os.path.abspath(__file__))
AREA      = 5.0
REPRO_DIR = os.path.join(BASE, "260603 New data set (reproducibility)")
OUT_DIR   = os.path.join(REPRO_DIR, "Final Plots", "PD plots")
os.makedirs(OUT_DIR, exist_ok=True)

# ── 260603 staging override for VC Polyol ──────────────────────────────────
VC_POLYOL_OVERRIDE = os.path.join(REPRO_DIR, "Overpotnital")

# Per-file staging override: <file stem> -> directory holding <dur>/<stem>_*.xlsx
# (staging layout has files directly in <dir>/<dur>/, no "Edited" subfolder).
# Keyed by the condition-specific stem so only that one panel is affected.
REPRO_ALT = os.path.join(REPRO_DIR, "Overpotential DATA raw alt")
DATA_OVERRIDE_DIRS = {
    "KB BM 1.5bp air": os.path.join(REPRO_ALT, "Air BP", "KB BM"),  # Air BP / KB BM only
}

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

# ── Durability style (I-V / PD curves) ─────────────────────────────────────
DUR_META = {
    "BOL": dict(ls="-",  marker="s", color="#000000"),
    "75K": dict(ls="--", marker="^", color="#444444"),
}
DUR_ORDER = ["BOL", "75K"]

# ── Voltage reference lines ────────────────────────────────────────────────
V_REF  = [0.70, 0.40]
V_CLRS = ["#C62828", "#1565C0"]   # dark-red, dark-blue

X_FMT = FuncFormatter(lambda x, p: f"{x:.1f}" if x >= 0.1 else f"{x:.2f}")
PANEL_FIGSIZE = (5.5, 4.0)


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


def compute_pd_ymax(iv_sample):
    pd_max = 0.0
    for cond_folder, _, _, k75 in ROWS:
        dur_dirs = {"BOL": "BOL", "75K": k75}
        fname = FNAMES.get((cond_folder, iv_sample))
        if not fname:
            continue
        for dur in DUR_ORDER:
            I, V = load_iv(cond_folder, dur_dirs[dur], fname, override_vc=True)
            if I is not None:
                pd_max = max(pd_max, ((I / 1000) * V).max())
    return np.ceil(pd_max / 0.2) * 0.2


# ── Per-panel drawing (axis labels intentionally omitted) ──────────────────
def draw_iv(ax, cond_folder, gas_str, bp_str, k75, fname, show_legend):
    dur_dirs = {"BOL": "BOL", "75K": k75}

    for dur in DUR_ORDER:
        I, V = load_iv(cond_folder, dur_dirs[dur], fname, override_vc=True) if fname else (None, None)
        if I is None:
            continue
        ax.plot(I / 1000, V, ls=DUR_META[dur]["ls"],
                color=DUR_META[dur]["color"], lw=1.6, zorder=3)
        P = I / 1000 * V
        idx_mp = int(np.argmax(P))
        ax.plot(I[idx_mp] / 1000, V[idx_mp],
                marker="D" if dur == "BOL" else "P",
                ms=9, zorder=7, linestyle="none",
                color=DUR_META[dur]["color"],
                markeredgecolor="white", markeredgewidth=1.2)

    for v_tgt, clr in zip(V_REF, V_CLRS):
        any_found = False
        for dur in DUR_ORDER:
            I, V = load_iv(cond_folder, dur_dirs[dur], fname, override_vc=True) if fname else (None, None)
            j = j_at_v(I, V, v_tgt)
            if not np.isnan(j):
                any_found = True
                ax.vlines(j / 1000, 0.2, v_tgt, colors=clr, lw=1.0, ls="--", zorder=4)
                ax.plot(j / 1000, v_tgt,
                        marker=DUR_META[dur]["marker"], ms=9,
                        color=clr, zorder=6, linestyle="none",
                        markeredgecolor="black", markeredgewidth=0.7)
        if any_found:
            ax.axhline(v_tgt, color=clr, lw=1.0, ls="--", zorder=4)
        ax.text(0.02, v_tgt - 0.025, f"{v_tgt:.1f} V",
                transform=ax.get_yaxis_transform(),
                color=clr, fontsize=FS - 2, va="top", ha="left")

    ax.set_xscale("log"); ax.set_xlim(0.005, 4); ax.set_ylim(0.2, 1.0)
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.2))
    ax.xaxis.set_major_formatter(X_FMT)
    ax.tick_params(direction="out", top=False, right=False, width=1.2, length=5)
    ax.tick_params(which="minor", width=0.8, length=3)

    ax.text(0.97, 0.97, f"{gas_str}\n{bp_str}",
            transform=ax.transAxes, ha="right", va="top",
            fontsize=FS, linespacing=1.45)

    if show_legend:
        dur_handles = [
            mlines.Line2D([], [], color=DUR_META[d]["color"],
                          ls=DUR_META[d]["ls"], lw=1.6, label=d)
            for d in DUR_ORDER
        ]
        ax.legend(handles=dur_handles,
                  loc="upper right", bbox_to_anchor=(0.76, 0.97),
                  fontsize=FS, handlelength=1.5, borderpad=0.2,
                  labelspacing=0.3, frameon=False)


def draw_pd(ax, cond_folder, gas_str, bp_str, k75, fname, pd_ymax,
            annotate=False, show_legend=False):
    dur_dirs = {"BOL": "BOL", "75K": k75}

    for dur in DUR_ORDER:
        I, V = load_iv(cond_folder, dur_dirs[dur], fname, override_vc=True) if fname else (None, None)
        if I is None:
            continue
        P = (I / 1000) * V
        ax.plot(I / 1000, P, ls=DUR_META[dur]["ls"],
                color=DUR_META[dur]["color"], lw=1.6, zorder=3)
        idx_mp = int(np.argmax(P))
        ax.plot(I[idx_mp] / 1000, P[idx_mp],
                marker="D" if dur == "BOL" else "P",
                ms=9, zorder=7, linestyle="none",
                color=DUR_META[dur]["color"],
                markeredgecolor="white", markeredgewidth=1.2)

    ax.set_xscale("log"); ax.set_xlim(0.005, 4); ax.set_ylim(0, pd_ymax)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=4, steps=[1, 2, 5, 10]))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda y, p: f"{y:.1f}"))
    ax.xaxis.set_major_formatter(X_FMT)
    ax.tick_params(direction="out", top=False, right=False, width=1.2, length=5)
    ax.tick_params(which="minor", width=0.8, length=3)

    if annotate:
        ax.text(0.03, 0.97, f"{gas_str}\n{bp_str}", transform=ax.transAxes,
                ha="left", va="top", fontsize=FS, linespacing=1.45)

    if show_legend:
        dur_handles = [
            mlines.Line2D([], [], color=DUR_META[d]["color"],
                          ls=DUR_META[d]["ls"], lw=1.6, label=d)
            for d in DUR_ORDER
        ]
        ax.legend(handles=dur_handles, loc="upper right",
                  fontsize=FS, handlelength=1.5, borderpad=0.2,
                  labelspacing=0.3, frameon=False)


# ── Figure-per-panel builders ───────────────────────────────────────────────
def make_panel(draw_fn, *args, **kwargs):
    fig = plt.figure(figsize=PANEL_FIGSIZE, dpi=300, constrained_layout=True)
    gs = GridSpec(1, 1, figure=fig)
    ax = fig.add_subplot(gs[0, 0])
    draw_fn(ax, *args, **kwargs)
    return fig


def save_panel(fig, name):
    out = os.path.join(OUT_DIR, name)
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {name}")


def split_pd_grid(iv_sample, prefix):
    """Splits PD_<sample>_grid.png (1 col x 4 rows) into 4 panels."""
    pd_ymax = compute_pd_ymax(iv_sample)
    for row_idx, (cond_folder, gas_str, bp_str, k75) in enumerate(ROWS):
        fname = FNAMES.get((cond_folder, iv_sample))
        fig = make_panel(draw_pd, cond_folder, gas_str, bp_str, k75, fname,
                          pd_ymax, annotate=True, show_legend=(row_idx == 0))
        save_panel(fig, f"{prefix}_grid_{COND_TAG[cond_folder]}.png")


def split_iv_pd_grid(iv_sample, prefix):
    """Splits IV_PD_<sample>_grid.png (2 cols x 4 rows) into 8 panels."""
    pd_ymax = compute_pd_ymax(iv_sample)
    for row_idx, (cond_folder, gas_str, bp_str, k75) in enumerate(ROWS):
        fname = FNAMES.get((cond_folder, iv_sample))
        tag = COND_TAG[cond_folder]

        fig = make_panel(draw_iv, cond_folder, gas_str, bp_str, k75, fname,
                          show_legend=(row_idx == 0))
        save_panel(fig, f"{prefix}_grid_{tag}_IV.png")

        fig = make_panel(draw_pd, cond_folder, gas_str, bp_str, k75, fname,
                          pd_ymax, annotate=False, show_legend=False)
        save_panel(fig, f"{prefix}_grid_{tag}_PD.png")


split_pd_grid("KB BM", "PD_KB_BM")

split_iv_pd_grid("CN BM",     "IV_PD_CN_BM")
split_iv_pd_grid("KB BM",     "IV_PD_KB_BM")
split_iv_pd_grid("VC Polyol", "IV_PD_VC_Polyol")
split_iv_pd_grid("AB Polyol", "IV_PD_AB_Polyol")
split_iv_pd_grid("CN Polyol", "IV_PD_CN_Polyol")
split_iv_pd_grid("KB Polyol", "IV_PD_KB_Polyol")
split_iv_pd_grid("VC BM",     "IV_PD_VC_BM")
split_iv_pd_grid("AB BM",     "IV_PD_AB_BM")

print("Done.")
