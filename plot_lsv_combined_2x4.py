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

# ---------------------------------------------------------------------------
# Style
# ---------------------------------------------------------------------------
plt.rcParams.update({
    "font.family":       "Arial",
    "font.size":         8,
    "axes.linewidth":    0.8,
    "xtick.major.width": 0.8, "ytick.major.width": 0.8,
    "xtick.minor.width": 0.5, "ytick.minor.width": 0.5,
    "xtick.direction":   "in", "ytick.direction": "in",
    "xtick.top":  True,  "ytick.right": True,
    "xtick.major.size":  4,   "ytick.major.size":  4,
    "xtick.minor.size":  2.5, "ytick.minor.size":  2.5,
    "legend.frameon":    True, "legend.framealpha": 0.9,
    "legend.edgecolor":  "none", "legend.fontsize": 7,
    "lines.linewidth":   1.4,
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
# File-name lookup  (cond_folder, sample) → stem
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

# ---------------------------------------------------------------------------
# Row order: Air 0 bar | Air 1.5 bar | O2 0 bar | O2 1.5 bar
# (cond_folder, gas_str, bp_str, dur75_subdir, xlim, pmax)
# ---------------------------------------------------------------------------
ROWS = [
    ("Air",    "H$_2$/Air",   "0 bar$_g$",   "75K",  1500, 450),
    ("Air BP", "H$_2$/Air",   "1.5 bar$_g$", "75K",  2500, 700),
    ("O2",     "H$_2$/O$_2$", "0 bar$_g$",   "75K",  2500, 750),
    ("O2 BP",  "H$_2$/O$_2$", "1.5 bar$_g$", "75k",  3500, 1000),
]

DURS = [("BOL", "BOL", "-"), ("30K", "30K", "--"), ("75K", None, ":")]

# ---------------------------------------------------------------------------
# Data loader  (reads from individual edited/unchanged xlsx files)
# ---------------------------------------------------------------------------
def load_iv(cond_folder, dur_dir, fname):
    # Only use the Edited/ subfolder — root .xlsx files are BOL copies, not real dur data
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

# ---------------------------------------------------------------------------
# Combined figure builder
# ---------------------------------------------------------------------------
def build_combined(sample_order, color_map, group_name):
    nrows = len(ROWS)
    fig, axes = plt.subplots(
        nrows, 2,
        figsize=(7.2, 3.2 * nrows),
        dpi=300,
        constrained_layout=True,
    )

    for row_idx, (cond_folder, gas_str, bp_str, k75, xlim, pmax) in enumerate(ROWS):
        ax_v = axes[row_idx, 0]
        ax_p = axes[row_idx, 1]

        durs = [("BOL", "BOL", "-"), ("30K", "30K", "--"), ("75K", k75, ":")]

        # Plot all samples × durabilities
        for samp in sample_order:
            clr   = color_map[samp]
            fname = FNAMES[(cond_folder, samp)]
            for _, dur_dir, ls in durs:
                I, V, P = load_iv(cond_folder, dur_dir, fname)
                if I is None:
                    print(f"  MISSING {cond_folder}/{dur_dir}/{fname}")
                    continue
                ax_v.plot(I, V, ls=ls, color=clr, lw=1.4)
                ax_p.plot(I, P, ls=ls, color=clr, lw=1.4)


        # Axes formatting
        xtick = 500 if xlim <= 2500 else 1000
        for ax in (ax_v, ax_p):
            ax.set_xlim(0, xlim)
            ax.tick_params(which="both", direction="out", top=False, right=False)
            ax.minorticks_on()
            ax.xaxis.set_major_locator(plt.MultipleLocator(xtick))
            ax.set_xlabel("Current Density (mAcm$^{-2}$)", fontsize=8)

        ax_v.set_ylim(0.2, 1.0)
        ax_v.set_ylabel("Cell Voltage (V)", fontsize=8)
        ax_v.yaxis.set_major_locator(plt.MultipleLocator(0.2))

        ptick = 200 if pmax >= 700 else 100
        ax_p.set_ylim(0, pmax)
        ax_p.set_ylabel("Power Density (mWcm$^{-2}$)", fontsize=8)
        ax_p.yaxis.set_major_locator(plt.MultipleLocator(ptick))

        # Condition annotation in all left (I-V) panels
        annot = (
            f"{gas_str}\nRH100\nPt 5 wt%\n"
            r"0.05 mg$_\mathregular{Pt}$/cm$^2$"
            f"\nIC 0.8, N212\n{bp_str}"
        )
        ax_v.text(0.97, 0.97, annot,
                  transform=ax_v.transAxes, ha="right", va="top",
                  fontsize=7, linespacing=1.45)

        # Legends — top row only
        if row_idx == 0:
            # Durability legend in top-left I-V panel
            dur_handles = [
                mlines.Line2D([], [], color="gray", ls="-",  lw=1.4, label="BOL"),
                mlines.Line2D([], [], color="gray", ls="--", lw=1.4, label="30K"),
                mlines.Line2D([], [], color="gray", ls=":",  lw=1.5, label="75K"),
            ]
            leg1 = ax_v.legend(handles=dur_handles, loc="lower left",
                               fontsize=7, handlelength=2.5,
                               borderpad=0.5, labelspacing=0.3)
            ax_v.add_artist(leg1)

            # Sample legend in top-right power panel
            samp_handles = [
                mlines.Line2D([], [], color=color_map[s], ls="-", lw=1.8, label=s)
                for s in sample_order
            ]
            ax_p.legend(handles=samp_handles, loc="upper right",
                        fontsize=7, handlelength=1.6,
                        borderpad=0.5, labelspacing=0.3)

    out_dir = os.path.join(BASE, "Durability I-V figure")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"LSV_combined_2x4_{group_name}.png")
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


# ---------------------------------------------------------------------------
# BOL-only combined figure  (all 8 samples, colour = support, ls = method)
# ---------------------------------------------------------------------------

# Colour by carbon support, line style by synthesis method
SUPPORT_CLR = {
    "CN": "#1f7a1f",   # dark green
    "KB": "#56b4e9",   # sky blue
    "VC": "#000000",   # black
    "AB": "#e69f00",   # amber
}

# (fname_key, color, linestyle, label)
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
    """One BOL-only 2×4 figure for a sample group. Legend + condition text
    appear only in the top-left (row 0, col 0) panel."""
    nrows = len(ROWS)
    fig, axes = plt.subplots(
        nrows, 2,
        figsize=(7.2, 3.2 * nrows),
        dpi=300,
        constrained_layout=True,
    )

    for row_idx, (cond_folder, gas_str, bp_str, k75, xlim, pmax) in enumerate(ROWS):
        ax_v = axes[row_idx, 0]
        ax_p = axes[row_idx, 1]

        for fname_key, clr, ls, lbl in sample_list:
            fname = FNAMES[(cond_folder, fname_key)]
            I, V, P = load_iv(cond_folder, "BOL", fname)
            if I is None:
                print(f"  MISSING BOL {cond_folder}/{fname}")
                continue
            ax_v.plot(I, V, ls=ls, color=clr, lw=1.4)
            ax_p.plot(I, P, ls=ls, color=clr, lw=1.4)

            # Dot at max-power point on I-V curve
            idx = np.argmax(P)
            ax_v.plot(I[idx], V[idx], marker="o", ms=4,
                      color=clr, zorder=5, linestyle="none")

        # Axes formatting
        xtick = 500 if xlim <= 2500 else 1000
        for ax in (ax_v, ax_p):
            ax.set_xlim(0, xlim)
            ax.tick_params(which="both", direction="out", top=False, right=False)
            ax.minorticks_on()
            ax.xaxis.set_major_locator(plt.MultipleLocator(xtick))
            ax.set_xlabel("Current Density (mAcm$^{-2}$)", fontsize=8)

        ax_v.set_ylim(0.2, 1.0)
        ax_v.set_ylabel("Cell Voltage (V)", fontsize=8)
        ax_v.yaxis.set_major_locator(plt.MultipleLocator(0.2))

        ptick = 200 if pmax >= 700 else 100
        ax_p.set_ylim(0, pmax)
        ax_p.set_ylabel("Power Density (mWcm$^{-2}$)", fontsize=8)
        ax_p.yaxis.set_major_locator(plt.MultipleLocator(ptick))

        # Legend only in top-left I-V panel
        if row_idx == 0:
            samp_handles = [
                mlines.Line2D([], [], color=clr, ls=ls, lw=1.4, label=lbl)
                for _, clr, ls, lbl in sample_list
            ]
            leg = ax_v.legend(handles=samp_handles, loc="lower left",
                              fontsize=7, handlelength=2.0,
                              borderpad=0.5, labelspacing=0.3)
            ax_v.add_artist(leg)

        # Condition annotation in all left (I-V) panels
        annot = (
            f"{gas_str}\nRH100\nPt 5 wt%\n"
            r"0.05 mg$_\mathregular{Pt}$/cm$^2$"
            f"\nIC 0.8, N212\n{bp_str}"
        )
        ax_v.text(0.97, 0.97, annot,
                  transform=ax_v.transAxes, ha="right", va="top",
                  fontsize=7, linespacing=1.45)

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
