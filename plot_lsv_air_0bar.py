import sys, os, openpyxl
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.lines as mlines

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

AREA = 5.0  # cm²
BASE = r"c:\Users\user\My Drive\KAIST MASc 2021\Laboratory Work\Protocol\overpotential calculation\For paper SI"

# ── matplotlib global style ───────────────────────────────────────────────────
plt.rcParams.update({
    "font.family":       "Arial",
    "font.size":         8,
    "axes.linewidth":    0.8,
    "xtick.major.width": 0.8,
    "ytick.major.width": 0.8,
    "xtick.minor.width": 0.6,
    "ytick.minor.width": 0.6,
    "xtick.direction":   "in",
    "ytick.direction":   "in",
    "xtick.top":         True,
    "ytick.right":       True,
    "xtick.major.size":  4,
    "ytick.major.size":  4,
    "xtick.minor.size":  2.5,
    "ytick.minor.size":  2.5,
    "axes.spines.top":   True,
    "axes.spines.right": True,
    "legend.frameon":    True,
    "legend.framealpha": 0.9,
    "legend.edgecolor":  "none",
    "legend.fontsize":   7,
    "lines.linewidth":   1.4,
})

# ── sample definitions ────────────────────────────────────────────────────────
MAIN = [
    ("CN BM",     "CN BM air",     "#1f7a1f"),   # dark green
    ("KB BM",     "KB BM air",     "#1a5fa8"),   # dark blue
    ("VC Polyol", "VC Polyol air", "#cc2200"),   # dark red
    ("AB Polyol", "AB Polyol air", "#d4a017"),   # gold
]
OTHER = [
    ("CN Polyol", "CN Polyol air", "#2dc7a0"),   # teal
    ("KB Polyol", "KB Polyol air", "#56b4e9"),   # sky blue
    ("VC BM",     "VC BM air",     "#1a1a1a"),   # near black
    ("AB BM",     "AB BM Air",     "#e69f00"),   # amber   ← capital A in filename
]

DURS = [
    ("BOL", "BOL", "-"),
    ("30K", "30K", "--"),
    ("75K", "75K", ":"),
]

ANNOT = (
    "H$_2$/Air\nRH100\nPt 5 wt%\n"
    "0.05 mg$_\\mathregular{Pt}$/cm$^2$\n"
    "IC 0.8, N212\n0 bar$_g$"
)


def load_iv(cond_folder, dur_dir, fname):
    folder = os.path.join(BASE, "Overpotnital", cond_folder, dur_dir)
    for sfx in ("_edited.xlsx", "_unchanged.xlsx", ".xlsx"):
        p = os.path.join(folder, fname + sfx)
        if os.path.exists(p):
            wb = openpyxl.load_workbook(p, data_only=True)
            ws = wb.active
            pts = [(ws.cell(r, 1).value, ws.cell(r, 2).value)
                   for r in range(4, ws.max_row + 1)
                   if ws.cell(r, 1).value is not None]
            wb.close()
            if not pts:
                return None, None, None
            I_d = np.array([a * 1000.0 / AREA for a, _ in pts])   # mA/cm²
            V   = np.array([v for _, v in pts])
            P   = I_d * V                                           # mW/cm²
            return I_d, V, P
    return None, None, None


def make_panel_pair(ax_v, ax_p, samples, cond_folder):
    """Plot I-V and I-P curves onto given axes for one sample group."""
    for samp_lbl, fname, color in samples:
        for dur_lbl, dur_dir, ls in DURS:
            I, V, P = load_iv(cond_folder, dur_dir, fname)
            if I is None:
                print(f"  MISSING: {cond_folder}/{dur_dir}/{fname}")
                continue
            ax_v.plot(I, V, ls=ls, color=color, lw=1.4)
            ax_p.plot(I, P, ls=ls, color=color, lw=1.4)


def format_axes(ax_v, ax_p, xlim, ylim_v, ylim_p):
    for ax in (ax_v, ax_p):
        ax.set_xlim(0, xlim)
        ax.set_xlabel("Current Density (mAcm$^{-2}$)", fontsize=9)
        ax.tick_params(which="both", direction="in", top=True, right=True)
        ax.minorticks_on()

    ax_v.set_ylim(*ylim_v)
    ax_v.set_ylabel("Cell Voltage (V)", fontsize=9)
    ax_v.yaxis.set_major_locator(plt.MultipleLocator(0.2))

    ax_p.set_ylim(*ylim_p)
    ax_p.set_ylabel("Power Density (mWcm$^{-2}$)", fontsize=9)
    ax_p.yaxis.set_major_locator(plt.MultipleLocator(100))

    ax_p.xaxis.set_major_locator(plt.MultipleLocator(500))
    ax_v.xaxis.set_major_locator(plt.MultipleLocator(500))


def add_legends(ax_v, ax_p, samples):
    # Durability legend → upper LEFT of voltage panel (matching reference)
    dur_handles = [
        mlines.Line2D([], [], color="gray", ls="-",  lw=1.4, label="BOL"),
        mlines.Line2D([], [], color="gray", ls="--", lw=1.4, label="30K"),
        mlines.Line2D([], [], color="gray", ls=":",  lw=1.5, label="75K"),
    ]
    leg1 = ax_v.legend(handles=dur_handles, loc="upper left",
                       fontsize=7, handlelength=2.5,
                       borderpad=0.5, labelspacing=0.3)
    ax_v.add_artist(leg1)

    # Sample color legend → upper RIGHT of power panel
    samp_handles = [
        mlines.Line2D([], [], color=c, ls="-", lw=1.8, label=lbl)
        for lbl, _, c in samples
    ]
    ax_p.legend(handles=samp_handles, loc="upper right",
                fontsize=7, handlelength=1.6,
                borderpad=0.5, labelspacing=0.3)


def add_annotation(ax_v, ax_p):
    # Annotation upper-right of voltage panel (clear of legend)
    ax_v.text(0.97, 0.98, ANNOT,
              transform=ax_v.transAxes, ha="right", va="top",
              fontsize=7, linespacing=1.45)
    # Annotation lower-right of power panel (below sample legend)
    ax_p.text(0.97, 0.30, ANNOT,
              transform=ax_p.transAxes, ha="right", va="top",
              fontsize=7, linespacing=1.45)


def build_figure(group_name, samples, cond_folder,
                 xlim=1600, ylim_v=(0.2, 1.0), ylim_p=(0, 450)):
    fig, (ax_v, ax_p) = plt.subplots(
        1, 2,
        figsize=(7.2, 3.4),
        dpi=300,
        constrained_layout=True,
    )

    make_panel_pair(ax_v, ax_p, samples, cond_folder)
    format_axes(ax_v, ax_p, xlim, ylim_v, ylim_p)
    add_legends(ax_v, ax_p, samples)
    add_annotation(ax_v, ax_p)

    out = os.path.join(BASE, f"LSV_Air0bar_{group_name}.png")
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


# ── run ───────────────────────────────────────────────────────────────────────
build_figure("Main",  MAIN,  "Air")
build_figure("Other", OTHER, "Air")
print("Done.")
