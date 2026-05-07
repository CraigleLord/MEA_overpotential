"""
Combined 2×4 LSV figure: 4 conditions (rows) × 2 panels (I-V | Power).
Produces two files: Main samples and Other samples.
"""

import os, sys, openpyxl
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.lines as mlines

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = os.path.dirname(os.path.abspath(__file__))
AREA = 5.0   # cm²

FS_TICK   = 22
FS_LABEL  = 22
FS_LEGEND = 17
FS_ANNOT  = 22

# ---------------------------------------------------------------------------
# Style
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Sample groups
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# File-name lookup
# ---------------------------------------------------------------------------
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
    ("Air",    "H$_2$/Air",   "0 bar$_g$",   "75K",  1500, 450),
    ("Air BP", "H$_2$/Air",   "1.5 bar$_g$", "75K",  2500, 700),
    ("O2",     "H$_2$/O$_2$", "0 bar$_g$",   "75K",  2500, 750),
    ("O2 BP",  "H$_2$/O$_2$", "1.5 bar$_g$", "75k",  3500, 1000),
]

XLIM_ALL = 3500
PMAX_ALL = 1000

# ---------------------------------------------------------------------------
# Data loader
# ---------------------------------------------------------------------------
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


def _apply_ax_style(ax_v, ax_p):
    for ax in (ax_v, ax_p):
        ax.set_xlim(0, XLIM_ALL)
        ax.minorticks_on()
        ax.xaxis.set_major_locator(plt.MultipleLocator(1000))
        ax.xaxis.set_minor_locator(plt.MultipleLocator(500))
    ax_v.set_ylim(0.2, 1.0)
    ax_v.yaxis.set_major_locator(plt.MultipleLocator(0.2))
    ax_p.set_ylim(0, PMAX_ALL)
    ax_p.yaxis.set_major_locator(plt.MultipleLocator(200))


def _add_shared_labels(fig, axes):
    """Single centered x-title, left y-title, left-of-col-1 power density title."""
    fig.supxlabel("Current Density (mAcm$^{-2}$)", fontsize=FS_LABEL, y=0.03)
    fig.supylabel("Cell Voltage (V)", fontsize=FS_LABEL)
    # Power Density label on left side of the 2nd column
    xl = axes[0, 1].get_position().x0
    fig.text(xl - 0.07, 0.5, "Power Density (mWcm$^{-2}$)",
             va="center", ha="center", rotation=90, fontsize=FS_LABEL)


# ---------------------------------------------------------------------------
# Combined figure (all durabilities)
# ---------------------------------------------------------------------------
def build_combined(sample_order, color_map, group_name):
    nrows = len(ROWS)
    fig, axes = plt.subplots(nrows, 2, figsize=(12, 4.0 * nrows), dpi=300)
    fig.subplots_adjust(hspace=0.30, wspace=0.35,
                        left=0.10, right=0.86, bottom=0.08, top=0.97)

    for row_idx, (cond_folder, gas_str, bp_str, k75, _, _) in enumerate(ROWS):
        ax_v = axes[row_idx, 0]
        ax_p = axes[row_idx, 1]
        durs = [("BOL", "BOL", "-"), ("30K", "30K", "--"), ("75K", k75, ":")]

        for samp in sample_order:
            clr   = color_map[samp]
            fname = FNAMES[(cond_folder, samp)]
            for _, dur_dir, ls in durs:
                I, V, P = load_iv(cond_folder, dur_dir, fname)
                if I is None:
                    print(f"  MISSING {cond_folder}/{dur_dir}/{fname}")
                    continue
                ax_v.plot(I, V, ls=ls, color=clr, lw=1.6)
                ax_p.plot(I, P, ls=ls, color=clr, lw=1.6)

        _apply_ax_style(ax_v, ax_p)

        annot = (
            f"{gas_str}\n"
            r"0.05 mg$_\mathregular{Pt}$/cm$^2$"
            f"\n{bp_str}"
        )
        ax_v.text(0.97, 0.97, annot,
                  transform=ax_v.transAxes, ha="right", va="top",
                  fontsize=FS_ANNOT, linespacing=1.45)

        if row_idx == 0:
            dur_handles = [
                mlines.Line2D([], [], color="gray", ls="-",  lw=1.6, label="BOL"),
                mlines.Line2D([], [], color="gray", ls="--", lw=1.6, label="30K"),
                mlines.Line2D([], [], color="gray", ls=":",  lw=1.6, label="75K"),
            ]
            ax_v.legend(handles=dur_handles, loc="lower right",
                        fontsize=FS_LEGEND, handlelength=2.5,
                        borderpad=0.5, labelspacing=0.3)

            samp_handles = [
                mlines.Line2D([], [], color=color_map[s], ls="-", lw=1.8, label=s)
                for s in sample_order
            ]
            ax_p.legend(handles=samp_handles, loc="lower right",
                        fontsize=FS_LEGEND, handlelength=1.6,
                        borderpad=0.5, labelspacing=0.3)

    _add_shared_labels(fig, axes)

    out_dir = os.path.join(BASE, "Durability I-V figure")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"LSV_combined_2x4_{group_name}.png")
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


# ---------------------------------------------------------------------------
# BOL-only figure
# ---------------------------------------------------------------------------
SUPPORT_CLR = {
    "CN": "#1f7a1f",
    "KB": "#56b4e9",
    "VC": "#000000",
    "AB": "#e69f00",
}

BOL_MAIN_SAMPLES = [
    ("CN BM",     SUPPORT_CLR["CN"], "-",  "CN BM"),
    ("KB BM",     SUPPORT_CLR["KB"], "-",  "KB BM"),
    ("VC Polyol", SUPPORT_CLR["VC"], "--", "VC Polyol"),
    ("AB Polyol", SUPPORT_CLR["AB"], "--", "AB Polyol"),
]
BOL_OTHER_SAMPLES = [
    ("CN Polyol", SUPPORT_CLR["CN"], "--", "CN Polyol"),
    ("KB Polyol", SUPPORT_CLR["KB"], "--", "KB Polyol"),
    ("VC BM",     SUPPORT_CLR["VC"], "-",  "VC BM"),
    ("AB BM",     SUPPORT_CLR["AB"], "-",  "AB BM"),
]


def build_bol_figure(sample_list, group_name):
    nrows = len(ROWS)
    fig, axes = plt.subplots(nrows, 2, figsize=(12, 4.0 * nrows), dpi=300)
    fig.subplots_adjust(hspace=0.30, wspace=0.35,
                        left=0.10, right=0.86, bottom=0.08, top=0.97)

    for row_idx, (cond_folder, gas_str, bp_str, k75, _, _) in enumerate(ROWS):
        ax_v = axes[row_idx, 0]
        ax_p = axes[row_idx, 1]

        for fname_key, clr, ls, lbl in sample_list:
            fname = FNAMES[(cond_folder, fname_key)]
            I, V, P = load_iv(cond_folder, "BOL", fname)
            if I is None:
                print(f"  MISSING BOL {cond_folder}/{fname}")
                continue
            ax_v.plot(I, V, ls=ls, color=clr, lw=1.6)
            ax_p.plot(I, P, ls=ls, color=clr, lw=1.6)
            idx = np.argmax(P)
            ax_v.plot(I[idx], V[idx], marker="o", ms=6,
                      color=clr, zorder=5, linestyle="none")

        _apply_ax_style(ax_v, ax_p)

        annot = (
            f"{gas_str}\n"
            r"0.05 mg$_\mathregular{Pt}$/cm$^2$"
            f"\n{bp_str}"
        )
        ax_v.text(0.97, 0.97, annot,
                  transform=ax_v.transAxes, ha="right", va="top",
                  fontsize=FS_ANNOT, linespacing=1.45)

        if row_idx == 0:
            samp_handles = [
                mlines.Line2D([], [], color=clr, ls=ls, lw=1.6, label=lbl)
                for _, clr, ls, lbl in sample_list
            ]
            ax_p.legend(handles=samp_handles, loc="lower right",
                        fontsize=FS_LEGEND, handlelength=2.0,
                        borderpad=0.5, labelspacing=0.3)

    _add_shared_labels(fig, axes)

    out_dir = os.path.join(BASE, "Durability I-V figure")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"LSV_BOL_{group_name}_2x4.png")
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------
build_combined(MAIN_ORDER,  MAIN_CLR,  "Main")
build_combined(OTHER_ORDER, OTHER_CLR, "Other")
build_bol_figure(BOL_MAIN_SAMPLES,  "Main")
build_bol_figure(BOL_OTHER_SAMPLES, "Other")
print("Done.")
