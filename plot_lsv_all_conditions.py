import sys, os, openpyxl
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.lines as mlines

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

AREA = 5.0
BASE = r"c:\Users\user\My Drive\KAIST MASc 2021\Laboratory Work\Protocol\overpotential calculation\For paper SI"

# ── global plot style ─────────────────────────────────────────────────────────
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

# ── colors  (VC = always black) ───────────────────────────────────────────────
MAIN_CLR = {
    "CN BM":     "#1f7a1f",   # dark green
    "KB BM":     "#56b4e9",   # sky blue
    "VC Polyol": "#000000",   # BLACK
    "AB Polyol": "#d4a017",   # gold
}
OTHER_CLR = {
    "CN Polyol": "#1f7a1f",   # same dark green as CN BM
    "KB Polyol": "#56b4e9",   # sky blue, same as KB BM
    "VC BM":     "#000000",   # BLACK
    "AB BM":     "#e69f00",   # amber
}
MAIN_ORDER  = ["CN BM",    "KB BM",    "VC Polyol", "AB Polyol"]
OTHER_ORDER = ["CN Polyol","KB Polyol","VC BM",     "AB BM"]

# ── file name lookup (cond_folder, sample) → stem ────────────────────────────
FNAMES = {
    ("O2",    "CN BM"):     "CN BM o2",
    ("O2",    "CN Polyol"): "CN Polyol O2",
    ("O2",    "KB BM"):     "KB BM O2",
    ("O2",    "KB Polyol"): "KB Polyol O2",
    ("O2",    "VC BM"):     "VC BM O2",
    ("O2",    "VC Polyol"): "VC Polyol O2",
    ("O2",    "AB BM"):     "AB BM  O2",     # double space
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
    ("Air",   "AB BM"):     "AB BM Air",     # capital A
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

# ── condition definitions ─────────────────────────────────────────────────────
# (label, cond_folder, gas_str, bp_str, dur_75K_subfolder, xlim, plim_max)
CONDITIONS = [
    ("O2_0bar",    "O2",     "H$_2$/O$_2$", "0 bar$_g$",   "75K",  2500, 750),
    ("O2_1.5bar",  "O2 BP",  "H$_2$/O$_2$", "1.5 bar$_g$", "75k",  3500, 1000),
    ("Air_0bar",   "Air",    "H$_2$/Air",   "0 bar$_g$",   "75K",  1500, 450),
    ("Air_1.5bar", "Air BP", "H$_2$/Air",   "1.5 bar$_g$", "75K",  2500, 700),
]

DURS = [("BOL", "BOL", "-"), ("30K", "30K", "--"), ("75K", None, ":")]
# 75K subfolder is filled per-condition below


def load_iv(cond_folder, dur_dir, fname):
    # Use Edited/ subfolder only — root .xlsx files are BOL copies
    folder = os.path.join(BASE, "Overpotnital", cond_folder, dur_dir, "Edited")
    for sfx in ("_edited.xlsx", "_unchanged.xlsx"):
        p = os.path.join(folder, fname + sfx)
        if os.path.exists(p):
            wb = openpyxl.load_workbook(p, data_only=True)
            ws  = wb.active
            pts = [(ws.cell(r,1).value, ws.cell(r,2).value)
                   for r in range(4, ws.max_row+1)
                   if isinstance(ws.cell(r,1).value, (int,float))
                   and isinstance(ws.cell(r,2).value, (int,float))]
            wb.close()
            if not pts:
                return None, None, None
            I_d = np.array([a * 1000.0 / AREA for a, _ in pts])
            V   = np.array([v for _, v in pts])
            P   = I_d * V
            return I_d, V, P
    return None, None, None


def build_figure(cond_label, cond_folder, gas_str, bp_str, k75,
                 xlim, plim_max, sample_order, color_map):

    durs = [("BOL","BOL","-"), ("30K","30K","--"), ("75K", k75, ":")]

    annot = (
        f"{gas_str}\nRH100\nPt 5 wt%\n"
        r"0.05 mg$_\mathregular{Pt}$/cm$^2$" + "\n"
        f"IC 0.8, N212\n{bp_str}"
    )

    group = "Main" if sample_order is MAIN_ORDER else "Other"
    fig, (ax_v, ax_p) = plt.subplots(
        1, 2, figsize=(7.2, 3.4), dpi=300, constrained_layout=True
    )

    for samp in sample_order:
        color = color_map[samp]
        fname = FNAMES[(cond_folder, samp)]
        for dur_lbl, dur_dir, ls in durs:
            I, V, P = load_iv(cond_folder, dur_dir, fname)
            if I is None:
                print(f"  MISSING: {cond_folder}/{dur_dir}/{fname}")
                continue
            ax_v.plot(I, V, ls=ls, color=color, lw=1.4)
            ax_p.plot(I, P, ls=ls, color=color, lw=1.4)

    # ── axes ─────────────────────────────────────────────────────────────────
    x_tick = 500 if xlim <= 2000 else 1000
    for ax in (ax_v, ax_p):
        ax.set_xlim(0, xlim)
        ax.set_xlabel("Current Density (mAcm$^{-2}$)", fontsize=9)
        ax.tick_params(which="both", direction="in", top=True, right=True)
        ax.minorticks_on()
        ax.xaxis.set_major_locator(plt.MultipleLocator(x_tick))

    ax_v.set_ylim(0.2, 1.0)
    ax_v.set_ylabel("Cell Voltage (V)", fontsize=9)
    ax_v.yaxis.set_major_locator(plt.MultipleLocator(0.2))

    p_tick = 200 if plim_max >= 800 else 100
    ax_p.set_ylim(0, plim_max)
    ax_p.set_ylabel("Power Density (mWcm$^{-2}$)", fontsize=9)
    ax_p.yaxis.set_major_locator(plt.MultipleLocator(p_tick))

    # ── legends ───────────────────────────────────────────────────────────────
    dur_handles = [
        mlines.Line2D([], [], color="gray", ls="-",  lw=1.4, label="BOL"),
        mlines.Line2D([], [], color="gray", ls="--", lw=1.4, label="30K"),
        mlines.Line2D([], [], color="gray", ls=":",  lw=1.5, label="75K"),
    ]
    leg1 = ax_v.legend(handles=dur_handles, loc="upper left",
                       fontsize=7, handlelength=2.5,
                       borderpad=0.5, labelspacing=0.3)
    ax_v.add_artist(leg1)

    samp_handles = [
        mlines.Line2D([], [], color=color_map[s], ls="-", lw=1.8, label=s)
        for s in sample_order
    ]
    ax_p.legend(handles=samp_handles, loc="upper right",
                fontsize=7, handlelength=1.6,
                borderpad=0.5, labelspacing=0.3)

    # ── annotation ────────────────────────────────────────────────────────────
    ax_v.text(0.97, 0.98, annot,
              transform=ax_v.transAxes, ha="right", va="top",
              fontsize=7, linespacing=1.45)
    ax_p.text(0.97, 0.30, annot,
              transform=ax_p.transAxes, ha="right", va="top",
              fontsize=7, linespacing=1.45)

    out_dir = os.path.join(BASE, "Durability I-V figure")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"LSV_{cond_label}_{group}.png")
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


# ── combined 4×2 grid (Air 0bar / Air 1.5bar / O2 0bar / O2 1.5bar) ──────────
GRID_CONDITIONS = [
    ("Air_0bar",   "Air",    "H$_2$/Air",   "0 bar$_g$",   "75K",  1500, 450),
    ("Air_1.5bar", "Air BP", "H$_2$/Air",   "1.5 bar$_g$", "75K",  2500, 700),
    ("O2_0bar",    "O2",     "H$_2$/O$_2$", "0 bar$_g$",   "75K",  2500, 750),
    ("O2_1.5bar",  "O2 BP",  "H$_2$/O$_2$", "1.5 bar$_g$", "75k",  3500, 1000),
]


def build_combined_grid(sample_order, color_map, group_name):
    nrows = len(GRID_CONDITIONS)
    fig, axes = plt.subplots(
        nrows, 2,
        figsize=(7.2, 3.2 * nrows),
        dpi=300,
        constrained_layout=True,
    )

    for row_idx, (_, cond_folder, gas_str, bp_str, k75, xlim, plim_max) in enumerate(GRID_CONDITIONS):
        ax_v = axes[row_idx, 0]
        ax_p = axes[row_idx, 1]

        durs = [("BOL", "BOL", "-"), ("30K", "30K", "--"), ("75K", k75, ":")]

        for samp in sample_order:
            color = color_map[samp]
            fname = FNAMES[(cond_folder, samp)]
            for _, dur_dir, ls in durs:
                I, V, P = load_iv(cond_folder, dur_dir, fname)
                if I is None:
                    print(f"  MISSING: {cond_folder}/{dur_dir}/{fname}")
                    continue
                ax_v.plot(I, V, ls=ls, color=color, lw=1.4)
                ax_p.plot(I, P, ls=ls, color=color, lw=1.4)

        # Axes formatting — uniform scale across all rows
        for ax in (ax_v, ax_p):
            ax.set_xlim(0, 3500)
            ax.set_xlabel("Current Density (mAcm$^{-2}$)", fontsize=8)
            ax.tick_params(which="both", direction="out", top=False, right=False)
            ax.minorticks_on()
            ax.xaxis.set_major_locator(plt.MultipleLocator(1000))

        ax_v.set_ylim(0.2, 1.0)
        ax_v.set_ylabel("Cell Voltage (V)", fontsize=8)
        ax_v.yaxis.set_major_locator(plt.MultipleLocator(0.2))

        ax_p.set_ylim(0, 1000)
        ax_p.set_ylabel("Power Density (mWcm$^{-2}$)", fontsize=8)
        ax_p.yaxis.set_major_locator(plt.MultipleLocator(200))

        # Condition annotation in all left (I-V) panels
        annot = (
            f"{gas_str}\nRH100\nPt 5 wt%\n"
            r"0.05 mg$_\mathregular{Pt}$/cm$^2$" + "\n"
            f"IC 0.8, N212\n{bp_str}"
        )
        ax_v.text(0.97, 0.97, annot,
                  transform=ax_v.transAxes, ha="right", va="top",
                  fontsize=7, linespacing=1.45)

        # Legends only in top-left panels
        if row_idx == 0:
            dur_handles = [
                mlines.Line2D([], [], color="gray", ls="-",  lw=1.4, label="BOL"),
                mlines.Line2D([], [], color="gray", ls="--", lw=1.4, label="30K"),
                mlines.Line2D([], [], color="gray", ls=":",  lw=1.5, label="75K"),
            ]
            leg1 = ax_v.legend(handles=dur_handles, loc="lower left",
                               fontsize=7, handlelength=2.5,
                               borderpad=0.5, labelspacing=0.3)
            ax_v.add_artist(leg1)

            samp_handles = [
                mlines.Line2D([], [], color=color_map[s], ls="-", lw=1.8, label=s)
                for s in sample_order
            ]
            ax_p.legend(handles=samp_handles, loc="upper right",
                        fontsize=7, handlelength=1.6,
                        borderpad=0.5, labelspacing=0.3)

    out_dir = os.path.join(BASE, "All figures important")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"LSV_durability_{group_name}.png")
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


# ── run all ───────────────────────────────────────────────────────────────────
build_combined_grid(MAIN_ORDER,  MAIN_CLR,  "Main")
build_combined_grid(OTHER_ORDER, OTHER_CLR, "Other")

print("All done.")
