"""
KB BM / KB Polyol: 4-row × 3-column figure.
  Col 0 : I-V curves for BOL/30K/75K overlaid; log10 x-axis;
           horizontal + vertical reference lines at 0.7 V and 0.4 V;
           durability-specific symbols at intersections.
  Col 1 : j @ 0.7 V bar chart — all group samples, waterfall BOL/30K/75K.
  Col 2 : j @ 0.4 V bar chart — same.
  Rows   : 4 test conditions (Air 0bp, Air 1.5bp, O2 0bp, O2 1.5bp).
Bar-chart y-axis in A cm⁻².
Output → 260603 New data set (reproducibility)/

Label-position notes
--------------------
Bar chart y-labels use ax.set_ylabel() on a thin invisible "label axes"
spanning each column's full height so they stay correctly placed regardless
of axis margins or font-size changes.
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

BASE    = os.path.dirname(os.path.abspath(__file__))
AREA    = 5.0
OUT_DIR = os.path.join(BASE, "260603 New data set (reproducibility)")
os.makedirs(OUT_DIR, exist_ok=True)

# ── 260603 staging override for VC Polyol ──────────────────────────────────
VC_POLYOL_OVERRIDE = os.path.join(OUT_DIR, "Overpotnital")

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

# ── Sample groups (matching plot_lsv_combined_2x4.py) ─────────────────────
MAIN_ORDER  = ["CN BM",     "KB BM",     "VC Polyol", "AB Polyol"]
OTHER_ORDER = ["CN Polyol", "KB Polyol", "VC BM",     "AB BM"]

MAIN_CLR  = {"CN BM":"#1f7a1f","KB BM":"#56b4e9","VC Polyol":"#000000","AB Polyol":"#d4a017"}
OTHER_CLR = {"CN Polyol":"#1f7a1f","KB Polyol":"#56b4e9","VC BM":"#000000","AB BM":"#e69f00"}

# ── File stems ─────────────────────────────────────────────────────────────
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

# ── Durability style (I-V column) ──────────────────────────────────────────
DUR_META = {
    "BOL": dict(ls="-",  marker="s", color="#000000"),
    "75K": dict(ls="--", marker="^", color="#444444"),
}
DUR_ORDER = ["BOL", "75K"]

# ── Bar chart style (matching build_combined) ──────────────────────────────
BAR_DUR_ORDER = ["75K", "BOL"]
BAR_STYLES = {
    "BOL": dict(facecolor="#cccccc", hatch="",    edgecolor="black", linewidth=0.5),
    "75K": dict(facecolor="#555555", hatch="",    edgecolor="black", linewidth=0.5),
}
BAR_LABELS = {"BOL": "0 K", "75K": "75 K"}

# ── Voltage reference lines ────────────────────────────────────────────────
V_REF  = [0.70, 0.40]
V_CLRS = ["#C62828", "#1565C0"]   # dark-red, dark-blue


# ── Helpers ────────────────────────────────────────────────────────────────
def load_iv(cond_folder, dur_dir, fname, override_vc=False):
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


# ── Pre-compute shared bar y-scale across BOTH figures ─────────────────────
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
    # Different rounding step per voltage so ymax is sensible for both scales
    steps = {V_REF[0]: 0.2, V_REF[1]: 0.5}   # 0.70V: small range; 0.40V: larger
    return {v: np.ceil(max(bar_j[v]) / steps[v]) * steps[v] if bar_j[v] else 2.0
            for v in V_REF}


SHARED_BAR_YMAXES = _collect_bar_j([MAIN_ORDER, OTHER_ORDER])


# ── Figure builder ─────────────────────────────────────────────────────────
def build_grid(iv_sample, sample_order, color_map, out_name, out_dir=None):
    nrows = len(ROWS)
    bar_ymaxes = SHARED_BAR_YMAXES   # identical scale for both figures

    # ── Layout: constrained_layout so ax.set_ylabel() never overlaps ticks ──
    fig = plt.figure(figsize=(5.5 * 3 + 2.5, 4.0 * nrows), dpi=300,
                     constrained_layout=True)
    gs = GridSpec(nrows, 3, figure=fig, width_ratios=[1, 1, 1])

    axes_iv  = [fig.add_subplot(gs[r, 0]) for r in range(nrows)]
    axes_07  = [fig.add_subplot(gs[r, 1]) for r in range(nrows)]
    axes_04  = [fig.add_subplot(gs[r, 2]) for r in range(nrows)]

    x_pos = np.arange(len(sample_order))

    for row_idx, (cond_folder, gas_str, bp_str, k75) in enumerate(ROWS):
        dur_dirs = {"BOL": "BOL", "30K": "30K", "75K": k75}
        fname_iv = FNAMES.get((cond_folder, iv_sample))

        # ── Col 0: I-V (log10 x-axis) ─────────────────────────────────────
        ax_iv = axes_iv[row_idx]

        for dur in DUR_ORDER:
            I, V = load_iv(cond_folder, dur_dirs[dur], fname_iv)
            if I is not None:
                ax_iv.plot(I / 1000, V, ls=DUR_META[dur]["ls"],
                           color=DUR_META[dur]["color"], lw=1.6, zorder=3)
                # Max power density — ◆ for BOL, ✚ (P) for 75K
                P = I / 1000 * V
                idx_mp = int(np.argmax(P))
                mp_marker = "D" if dur == "BOL" else "P"
                ax_iv.plot(I[idx_mp] / 1000, V[idx_mp],
                           marker=mp_marker, ms=9, zorder=7, linestyle="none",
                           color=DUR_META[dur]["color"],
                           markeredgecolor="white", markeredgewidth=1.2)

        for v_tgt, clr in zip(V_REF, V_CLRS):
            any_found = False
            for dur in DUR_ORDER:
                I, V = load_iv(cond_folder, dur_dirs[dur], fname_iv)
                j = j_at_v(I, V, v_tgt)
                if not np.isnan(j):
                    any_found = True
                    ax_iv.vlines(j / 1000, 0.2, v_tgt,
                                 colors=clr, lw=1.0, ls="--", zorder=4)
                    ax_iv.plot(j / 1000, v_tgt,
                               marker=DUR_META[dur]["marker"], ms=9,
                               color=clr, zorder=6, linestyle="none",
                               markeredgecolor="black", markeredgewidth=0.7)
            if any_found:
                ax_iv.axhline(v_tgt, color=clr, lw=1.0, ls="--", zorder=4)

        ax_iv.set_xscale("log")
        ax_iv.set_xlim(0.005, 4)     # A cm⁻²
        ax_iv.set_ylim(0.2, 1.0)
        ax_iv.yaxis.set_major_locator(plt.MultipleLocator(0.2))
        import matplotlib.ticker as ticker
        ax_iv.xaxis.set_major_formatter(
            ticker.FuncFormatter(lambda x, p: f"{x:.1f}" if x >= 0.1 else f"{x:.2f}")
        )
        ax_iv.tick_params(direction="out", top=False, right=False,
                          width=1.2, length=5)
        ax_iv.tick_params(which="minor", width=0.8, length=3)

        # Condition annotation (no Pt loading line)
        annot = f"{gas_str}\n{bp_str}"
        ax_iv.text(0.97, 0.97, annot, transform=ax_iv.transAxes,
                   ha="right", va="top", fontsize=FS, linespacing=1.45)

        # Voltage labels below each reference line, left edge
        for v_tgt, clr in zip(V_REF, V_CLRS):
            ax_iv.text(0.02, v_tgt - 0.025, f"{v_tgt:.1f} V",
                       transform=ax_iv.get_yaxis_transform(),
                       color=clr, fontsize=FS - 2, va="top", ha="left")

        if row_idx == 0:
            ax_iv.set_title("I–V", fontsize=FS, fontweight="bold")

            dur_handles = [
                mlines.Line2D([], [], color=DUR_META[d]["color"],
                              ls=DUR_META[d]["ls"], lw=1.6, label=d)
                for d in DUR_ORDER
            ]
            # Legend immediately left of condition text, same top (y=0.97)
            ax_iv.legend(handles=dur_handles,
                         loc="upper right",
                         bbox_to_anchor=(0.76, 0.97),
                         fontsize=FS,
                         handlelength=1.5, borderpad=0.2,
                         labelspacing=0.3, frameon=False)

        # ── Bar chart columns ──────────────────────────────────────────────
        for ax_b, v_tgt, v_clr in zip([axes_07[row_idx], axes_04[row_idx]],
                                       V_REF, V_CLRS):
            i_abs = {}
            for dur in DUR_ORDER:
                vals = []
                for samp in sample_order:
                    fname = FNAMES.get((cond_folder, samp))
                    I, V  = load_iv(cond_folder, dur_dirs[dur], fname,
                                    override_vc=True) if fname else (None, None)
                    j = j_at_v(I, V, v_tgt)
                    vals.append(0.0 if np.isnan(j) else j / 1000)  # A cm⁻²
                i_abs[dur] = np.array(vals)

            increments = {
                "75K": i_abs["75K"],
                "BOL": np.maximum(i_abs["BOL"] - i_abs["75K"], 0),
            }

            bottoms = np.zeros(len(sample_order))
            for dur_key in BAR_DUR_ORDER:
                ax_b.bar(x_pos, increments[dur_key], bottom=bottoms,
                         width=0.55, label=BAR_LABELS[dur_key],
                         **BAR_STYLES[dur_key])
                bottoms += increments[dur_key]

            # Durability markers only on the iv_sample bar
            if iv_sample in sample_order:
                xi_main = sample_order.index(iv_sample)
                for dur in DUR_ORDER:
                    j_val = i_abs[dur][xi_main]
                    if j_val > 0:
                        ax_b.plot(xi_main, j_val,
                                  marker=DUR_META[dur]["marker"], ms=10,
                                  color=v_clr, zorder=8, linestyle="none",
                                  markeredgecolor="black", markeredgewidth=0.7)

            ax_b.set_xlim(-0.5, len(sample_order) - 0.5)
            ax_b.set_ylim(0, bar_ymaxes[v_tgt])
            from matplotlib.ticker import MaxNLocator
            ax_b.yaxis.set_major_locator(MaxNLocator(nbins=4, steps=[1,2,5,10]))
            ax_b.yaxis.set_major_formatter(
                ticker.FuncFormatter(lambda y, p: f"{y:.1f}")
            )
            ax_b.set_xticks(x_pos)
            if row_idx == nrows - 1:
                short = [s.replace(" ", "\n") for s in sample_order]
                ax_b.set_xticklabels(short, fontsize=FS, rotation=0, ha="center")
            else:
                ax_b.set_xticklabels([])
            ax_b.tick_params(axis="y", labelsize=FS)
            ax_b.tick_params(axis="x", which="minor", bottom=False)
            ax_b.minorticks_on()
            ax_b.tick_params(direction="out", top=False, right=False,
                             width=1.2, length=5)

            if row_idx == 0:
                ax_b.set_title(f"j @ {v_tgt:.1f} V",
                               fontsize=FS, fontweight="bold", color=v_clr)
                if v_tgt == V_REF[0]:
                    leg_handles = [
                        Patch(label=BAR_LABELS[d], **BAR_STYLES[d])
                        for d in ["BOL", "75K"]
                    ]
                    ax_b.legend(handles=leg_handles, loc="upper right",
                                fontsize=FS, frameon=False,
                                handlelength=1.2, labelspacing=0.3)

    # ── Shared axis labels ─────────────────────────────────────────────────
    # Cell Voltage y-label (I-V column)
    fig.supylabel("Cell Voltage (V)", fontsize=FS)

    # Current Density x-label: put on the bottom I-V subplot only so
    # constrained_layout handles the gap automatically — no overlap.
    axes_iv[-1].set_xlabel("Current Density (A cm$^{-2}$)", fontsize=FS)

    # Bar-chart y-labels: one per column, vertically centred.
    # Step 1: set ylabel on ALL rows so constrained_layout reserves the space.
    for ax in axes_07:
        ax.set_ylabel("Current Density (A cm$^{-2}$)", fontsize=FS)
    for ax in axes_04:
        ax.set_ylabel("Current Density (A cm$^{-2}$)", fontsize=FS)

    # Step 2: draw so positions are finalised, then freeze layout.
    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    fig.set_layout_engine("none")   # freeze — no re-layout on save

    # Step 3: for each bar column, hide all row ylabels and place one centred.
    for col_axes in [axes_07, axes_04]:
        # Vertical centre of the column in figure fraction coords
        y_top = col_axes[0].get_position().y1
        y_bot = col_axes[-1].get_position().y0
        y_c   = (y_top + y_bot) / 2

        # X position of the ylabel text (same for all rows in the column)
        lbl_bb = col_axes[0].yaxis.label.get_window_extent(renderer)
        fig_w  = fig.get_figwidth() * fig.dpi
        x_c    = (lbl_bb.x0 + lbl_bb.x1) / 2 / fig_w

        # Hide all per-row ylabels
        for ax in col_axes:
            ax.yaxis.label.set_visible(False)

        # Place one centred label
        fig.text(x_c, y_c, "Current Density (A cm$^{-2}$)",
                 ha="center", va="center", rotation=90, fontsize=FS)

    out = os.path.join(out_dir or OUT_DIR, out_name)
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


def build_pd_grid(iv_sample, out_name):
    """4-row × 1-column power density curves for iv_sample (BOL + 75K)."""
    nrows = len(ROWS)

    # Pre-compute global y-max and x-max across all conditions and durabilities
    pd_max = 0.0
    x_global_max = 0.0
    for cond_folder, _, _, k75 in ROWS:
        dur_dirs = {"BOL": "BOL", "75K": k75}
        fname = FNAMES.get((cond_folder, iv_sample))
        if not fname:
            continue
        for dur in DUR_ORDER:
            I, V = load_iv(cond_folder, dur_dirs[dur], fname)
            if I is not None:
                P = (I / 1000) * V   # W/cm²
                pd_max = max(pd_max, P.max())
                x_global_max = max(x_global_max, (I / 1000).max())
    ymax = np.ceil(pd_max / 0.2) * 0.2
    xlim = 4.0   # exactly matches the I-V subplot x-axis in build_grid

    fig = plt.figure(figsize=(5.5, 4.0 * nrows), dpi=300,
                     constrained_layout=True)
    gs = GridSpec(nrows, 1, figure=fig)
    axes = [fig.add_subplot(gs[r, 0]) for r in range(nrows)]

    for row_idx, (cond_folder, gas_str, bp_str, k75) in enumerate(ROWS):
        ax = axes[row_idx]
        dur_dirs = {"BOL": "BOL", "75K": k75}
        fname = FNAMES.get((cond_folder, iv_sample))

        for dur in DUR_ORDER:
            I, V = load_iv(cond_folder, dur_dirs[dur], fname) if fname else (None, None)
            if I is None:
                continue
            P = (I / 1000) * V   # W/cm²
            ax.plot(I / 1000, P,
                    ls=DUR_META[dur]["ls"],
                    color=DUR_META[dur]["color"], lw=1.6, zorder=3)
            # Peak power marker
            idx_mp = int(np.argmax(P))
            mp_marker = "D" if dur == "BOL" else "P"
            ax.plot(I[idx_mp] / 1000, P[idx_mp],
                    marker=mp_marker, ms=9, zorder=7, linestyle="none",
                    color=DUR_META[dur]["color"],
                    markeredgecolor="white", markeredgewidth=1.2)

        from matplotlib.ticker import MaxNLocator, FuncFormatter
        ax.set_xscale("log")
        ax.set_xlim(0.005, xlim)
        ax.set_ylim(0, ymax)
        ax.xaxis.set_major_formatter(
            FuncFormatter(lambda x, p: f"{x:.1f}" if x >= 0.1 else f"{x:.2f}")
        )
        ax.yaxis.set_major_locator(MaxNLocator(nbins=4, steps=[1,2,5,10]))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda y, p: f"{y:.1f}"))
        ax.tick_params(direction="out", top=False, right=False,
                       width=1.2, length=5)
        ax.tick_params(which="minor", width=0.8, length=3)

        annot = f"{gas_str}\n{bp_str}"
        ax.text(0.03, 0.97, annot, transform=ax.transAxes,
                ha="left", va="top", fontsize=FS, linespacing=1.45)

        if row_idx == 0:
            dur_handles = [
                mlines.Line2D([], [], color=DUR_META[d]["color"],
                              ls=DUR_META[d]["ls"], lw=1.6, label=d)
                for d in DUR_ORDER
            ]
            ax.legend(handles=dur_handles,
                      loc="upper right",
                      fontsize=FS, handlelength=1.5, borderpad=0.2,
                      labelspacing=0.3, frameon=False)

    fig.supylabel("Power Density (W cm$^{-2}$)", fontsize=FS)
    axes[-1].set_xlabel("Current Density (A cm$^{-2}$)", fontsize=FS)

    out = os.path.join(OUT_DIR, out_name)
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


def build_iv_pd_grid(iv_sample, out_name):
    """4-row × 2-column: I-V (col 0) + Power Density (col 1)."""
    from matplotlib.ticker import MaxNLocator, FuncFormatter
    nrows = len(ROWS)

    # Global PD y-max
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
    pd_ymax = np.ceil(pd_max / 0.2) * 0.2

    fig = plt.figure(figsize=(5.5 * 2 + 1, 4.0 * nrows), dpi=300,
                     constrained_layout=True)
    gs  = GridSpec(nrows, 2, figure=fig)
    axes_iv = [fig.add_subplot(gs[r, 0]) for r in range(nrows)]
    axes_pd = [fig.add_subplot(gs[r, 1]) for r in range(nrows)]

    x_fmt = FuncFormatter(lambda x, p: f"{x:.1f}" if x >= 0.1 else f"{x:.2f}")

    for row_idx, (cond_folder, gas_str, bp_str, k75) in enumerate(ROWS):
        dur_dirs = {"BOL": "BOL", "75K": k75}
        fname    = FNAMES.get((cond_folder, iv_sample))
        ax_iv    = axes_iv[row_idx]
        ax_pd    = axes_pd[row_idx]

        # ── Col 0: I-V (exact copy of build_grid col 0) ───────────────────
        for dur in DUR_ORDER:
            I, V = load_iv(cond_folder, dur_dirs[dur], fname, override_vc=True) if fname else (None, None)
            if I is None:
                continue
            ax_iv.plot(I / 1000, V, ls=DUR_META[dur]["ls"],
                       color=DUR_META[dur]["color"], lw=1.6, zorder=3)
            P = I / 1000 * V
            idx_mp = int(np.argmax(P))
            ax_iv.plot(I[idx_mp] / 1000, V[idx_mp],
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
                    ax_iv.vlines(j / 1000, 0.2, v_tgt, colors=clr, lw=1.0, ls="--", zorder=4)
                    ax_iv.plot(j / 1000, v_tgt,
                               marker=DUR_META[dur]["marker"], ms=9,
                               color=clr, zorder=6, linestyle="none",
                               markeredgecolor="black", markeredgewidth=0.7)
            if any_found:
                ax_iv.axhline(v_tgt, color=clr, lw=1.0, ls="--", zorder=4)
            ax_iv.text(0.02, v_tgt - 0.025, f"{v_tgt:.1f} V",
                       transform=ax_iv.get_yaxis_transform(),
                       color=clr, fontsize=FS - 2, va="top", ha="left")

        ax_iv.set_xscale("log"); ax_iv.set_xlim(0.005, 4); ax_iv.set_ylim(0.2, 1.0)
        ax_iv.yaxis.set_major_locator(plt.MultipleLocator(0.2))
        ax_iv.xaxis.set_major_formatter(x_fmt)
        ax_iv.tick_params(direction="out", top=False, right=False, width=1.2, length=5)
        ax_iv.tick_params(which="minor", width=0.8, length=3)

        ax_iv.text(0.97, 0.97, f"{gas_str}\n{bp_str}",
                   transform=ax_iv.transAxes, ha="right", va="top",
                   fontsize=FS, linespacing=1.45)

        if row_idx == 0:
            dur_handles = [
                mlines.Line2D([], [], color=DUR_META[d]["color"],
                              ls=DUR_META[d]["ls"], lw=1.6, label=d)
                for d in DUR_ORDER
            ]
            ax_iv.legend(handles=dur_handles,
                         loc="upper right", bbox_to_anchor=(0.76, 0.97),
                         fontsize=FS, handlelength=1.5, borderpad=0.2,
                         labelspacing=0.3, frameon=False)

        # ── Col 1: Power Density (exact copy of build_pd_grid) ────────────
        for dur in DUR_ORDER:
            I, V = load_iv(cond_folder, dur_dirs[dur], fname, override_vc=True) if fname else (None, None)
            if I is None:
                continue
            P = (I / 1000) * V
            ax_pd.plot(I / 1000, P, ls=DUR_META[dur]["ls"],
                       color=DUR_META[dur]["color"], lw=1.6, zorder=3)
            idx_mp = int(np.argmax(P))
            ax_pd.plot(I[idx_mp] / 1000, P[idx_mp],
                       marker="D" if dur == "BOL" else "P",
                       ms=9, zorder=7, linestyle="none",
                       color=DUR_META[dur]["color"],
                       markeredgecolor="white", markeredgewidth=1.2)

        ax_pd.set_xscale("log"); ax_pd.set_xlim(0.005, 4); ax_pd.set_ylim(0, pd_ymax)
        ax_pd.yaxis.set_major_locator(MaxNLocator(nbins=4, steps=[1, 2, 5, 10]))
        ax_pd.yaxis.set_major_formatter(FuncFormatter(lambda y, p: f"{y:.1f}"))
        ax_pd.xaxis.set_major_formatter(x_fmt)
        ax_pd.tick_params(direction="out", top=False, right=False, width=1.2, length=5)
        ax_pd.tick_params(which="minor", width=0.8, length=3)

    # ── Shared labels using freeze-and-centre approach ─────────────────────
    for ax in axes_iv:
        ax.set_ylabel("Cell Voltage (V)", fontsize=FS)
    for ax in axes_pd:
        ax.set_ylabel("Power Density (W cm$^{-2}$)", fontsize=FS)
    axes_iv[-1].set_xlabel("Current Density (A cm$^{-2}$)", fontsize=FS)
    axes_pd[-1].set_xlabel("Current Density (A cm$^{-2}$)", fontsize=FS)

    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    fig.set_layout_engine("none")

    for col_axes, label in [(axes_iv, "Cell Voltage (V)"),
                            (axes_pd, "Power Density (W cm$^{-2}$)")]:
        y_c = (col_axes[0].get_position().y1 + col_axes[-1].get_position().y0) / 2
        lbl_bb = col_axes[0].yaxis.label.get_window_extent(renderer)
        x_c = (lbl_bb.x0 + lbl_bb.x1) / 2 / (fig.get_figwidth() * fig.dpi)
        for ax in col_axes:
            ax.yaxis.label.set_visible(False)
        fig.text(x_c, y_c, label, ha="center", va="center",
                 rotation=90, fontsize=FS)

    out = os.path.join(OUT_DIR, out_name)
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


def build_bar_only_grid(sample_order, out_name, out_dir=None):
    """4-row × 2-column bar charts only: j@0.7V (col 0) and j@0.4V (col 1).
    Identical bar style to build_grid; uses SHARED_BAR_YMAXES."""
    from matplotlib.ticker import MaxNLocator, FuncFormatter
    nrows = len(ROWS)

    fig = plt.figure(figsize=(5.5 * 2 + 1, 4.0 * nrows), dpi=300,
                     constrained_layout=True)
    gs      = GridSpec(nrows, 2, figure=fig)
    axes_07 = [fig.add_subplot(gs[r, 0]) for r in range(nrows)]
    axes_04 = [fig.add_subplot(gs[r, 1]) for r in range(nrows)]

    x_pos = np.arange(len(sample_order))

    for row_idx, (cond_folder, gas_str, bp_str, k75) in enumerate(ROWS):
        dur_dirs = {"BOL": "BOL", "30K": "30K", "75K": k75}

        for ax_b, v_tgt, v_clr in zip([axes_07[row_idx], axes_04[row_idx]],
                                       V_REF, V_CLRS):
            # Compute j values (A/cm²) for each sample and durability
            i_abs = {}
            for dur in DUR_ORDER:
                vals = []
                for samp in sample_order:
                    fname = FNAMES.get((cond_folder, samp))
                    I, V  = load_iv(cond_folder, dur_dirs[dur], fname,
                                    override_vc=True) if fname else (None, None)
                    j = j_at_v(I, V, v_tgt)
                    vals.append((0.0 if np.isnan(j) else j) / 1000)   # A/cm²
                i_abs[dur] = np.array(vals)

            increments = {
                "75K": i_abs["75K"],
                "BOL": np.maximum(i_abs["BOL"] - i_abs["75K"], 0),
            }

            bottoms = np.zeros(len(sample_order))
            for dur_key in BAR_DUR_ORDER:
                ax_b.bar(x_pos, increments[dur_key], bottom=bottoms,
                         width=0.55, label=BAR_LABELS[dur_key],
                         **BAR_STYLES[dur_key])
                bottoms += increments[dur_key]

            ax_b.set_xlim(-0.5, len(sample_order) - 0.5)
            ax_b.set_ylim(0, SHARED_BAR_YMAXES[v_tgt])
            ax_b.yaxis.set_major_locator(MaxNLocator(nbins=4, steps=[1, 2, 5, 10]))
            ax_b.yaxis.set_major_formatter(FuncFormatter(lambda y, p: f"{y:.1f}"))
            ax_b.set_xticks(x_pos)
            if row_idx == nrows - 1:
                short = [s.replace(" ", "\n") for s in sample_order]
                ax_b.set_xticklabels(short, fontsize=FS, rotation=0, ha="center")
            else:
                ax_b.set_xticklabels([])
            ax_b.tick_params(axis="y", labelsize=FS)
            ax_b.tick_params(axis="x", which="minor", bottom=False)
            ax_b.minorticks_on()
            ax_b.tick_params(direction="out", top=False, right=False,
                             width=1.2, length=5)

            if row_idx == 0:
                ax_b.set_title(f"j @ {v_tgt:.1f} V",
                               fontsize=FS, fontweight="bold", color=v_clr)
                if v_tgt == V_REF[0]:
                    leg_handles = [Patch(label=BAR_LABELS[d], **BAR_STYLES[d])
                                   for d in ["BOL", "75K"]]
                    ax_b.legend(handles=leg_handles, loc="upper right",
                                fontsize=FS, frameon=False,
                                handlelength=1.2, labelspacing=0.3)

            # Condition label: left column only, top-left
            if v_tgt == V_REF[0]:
                ax_b.text(0.03, 0.97, f"{gas_str}\n{bp_str}",
                          transform=ax_b.transAxes,
                          ha="left", va="top", fontsize=FS, linespacing=1.45)

    # ── Shared y-labels (freeze-and-centre) ────────────────────────────────
    for ax in axes_07:
        ax.set_ylabel("Current Density (A cm$^{-2}$)", fontsize=FS)
    for ax in axes_04:
        ax.set_ylabel("Current Density (A cm$^{-2}$)", fontsize=FS)

    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    fig.set_layout_engine("none")

    for col_axes in [axes_07, axes_04]:
        y_c    = (col_axes[0].get_position().y1 + col_axes[-1].get_position().y0) / 2
        lbl_bb = col_axes[0].yaxis.label.get_window_extent(renderer)
        x_c    = (lbl_bb.x0 + lbl_bb.x1) / 2 / (fig.get_figwidth() * fig.dpi)
        for ax in col_axes:
            ax.yaxis.label.set_visible(False)
        fig.text(x_c, y_c, "Current Density (A cm$^{-2}$)",
                 ha="center", va="center", rotation=90, fontsize=FS)

    out = os.path.join(out_dir or OUT_DIR, out_name)
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


build_grid("KB BM",     MAIN_ORDER,  MAIN_CLR,  "LSV_KB_BM_dur_grid.png")
build_grid("KB Polyol", OTHER_ORDER, OTHER_CLR, "LSV_KB_Polyol_dur_grid.png")
# Both figures use SHARED_BAR_YMAXES so bar chart scales are identical.

build_pd_grid("KB BM", "PD_KB_BM_grid.png")
build_iv_pd_grid("CN BM",     "IV_PD_CN_BM_grid.png")
build_iv_pd_grid("KB BM",     "IV_PD_KB_BM_grid.png")
build_iv_pd_grid("VC Polyol", "IV_PD_VC_Polyol_grid.png")
build_iv_pd_grid("AB Polyol", "IV_PD_AB_Polyol_grid.png")
build_iv_pd_grid("CN Polyol", "IV_PD_CN_Polyol_grid.png")
build_iv_pd_grid("KB Polyol", "IV_PD_KB_Polyol_grid.png")
build_iv_pd_grid("VC BM",     "IV_PD_VC_BM_grid.png")
build_iv_pd_grid("AB BM",     "IV_PD_AB_BM_grid.png")
FINAL_PLOTS = os.path.join(OUT_DIR, "Final Plots")
os.makedirs(FINAL_PLOTS, exist_ok=True)
build_bar_only_grid(OTHER_ORDER, "Bar_Other_0p7_0p4V_grid.png", out_dir=FINAL_PLOTS)
print("Done.")
