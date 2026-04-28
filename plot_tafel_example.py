import sys, os, openpyxl
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy import stats

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

AREA  = 5.0    # cm²
E_EQ  = 1.23   # V  (equilibrium potential)
BASE  = r"c:\Users\user\My Drive\KAIST MASc 2021\Laboratory Work\Protocol\overpotential calculation\For paper SI"
OUT   = os.path.join(BASE, "Durability I-V figure")

plt.rcParams.update({
    "font.family":       "Arial",
    "font.size":         9,
    "axes.linewidth":    0.8,
    "xtick.major.width": 0.8, "ytick.major.width": 0.8,
    "xtick.direction":   "in", "ytick.direction": "in",
    "xtick.top":  True,  "ytick.right": True,
    "xtick.major.size":  4, "ytick.major.size": 4,
})

# ── load data ─────────────────────────────────────────────────────────────────
def load_iv(cond_folder, dur_dir, fname):
    folder = os.path.join(BASE, "Overpotnital", cond_folder, dur_dir)
    for sfx in ("_edited.xlsx", "_unchanged.xlsx", ".xlsx"):
        p = os.path.join(folder, fname + sfx)
        if os.path.exists(p):
            wb = openpyxl.load_workbook(p, data_only=True)
            ws  = wb.active
            pts = [(ws.cell(r,1).value, ws.cell(r,2).value)
                   for r in range(4, ws.max_row+1)
                   if ws.cell(r,1).value is not None]
            wb.close()
            I_A = np.array([a for a, _ in pts])
            V   = np.array([v for _, v in pts])
            j   = I_A / AREA       # A/cm²
            return j, V
    raise FileNotFoundError(fname)

# ── Tafel fit ─────────────────────────────────────────────────────────────────
def tafel_fit(j, eta, ln_lo=-3.5, ln_hi=-2.5):
    """Linear regression in ln|j| ∈ [ln_lo, ln_hi]."""
    ln_j = np.log(j)
    mask = (ln_j >= ln_lo) & (ln_j <= ln_hi)
    slope, intercept, r, p, se = stats.linregress(ln_j[mask], eta[mask])
    return slope, intercept, ln_j, mask

# ── plot ──────────────────────────────────────────────────────────────────────
def plot_tafel(cond_folder, dur_dir, fname, title, outname,
               xlim=(-5, 1), ylim=(-1.0, -0.3)):

    j, V = load_iv(cond_folder, dur_dir, fname)
    eta  = V - E_EQ

    slope, intercept, ln_j, mask = tafel_fit(j, eta)
    print(f"{title}")
    print(f"  Tafel region points: {mask.sum()}")
    print(f"  Slope:     {slope:.4f}")
    print(f"  Intercept: {intercept:.3f}")

    # regression line extrapolated across xlim
    x_line = np.linspace(xlim[0], xlim[1], 200)
    y_line = slope * x_line + intercept

    fig, ax = plt.subplots(figsize=(4.5, 4.0), dpi=300)

    # all data — filled black squares
    ax.plot(ln_j, eta, "ks", markersize=2.5, markeredgewidth=0, zorder=3)

    # Tafel regression — red dashed line
    ax.plot(x_line, y_line, "r--", lw=1.4, zorder=4)

    # annotation
    ann_txt = f"Intercept: {intercept:.3f}\nSlope: {slope:.4f}"
    ax.text(0.05, 0.25, ann_txt,
            transform=ax.transAxes, ha="left", va="top",
            fontsize=9, linespacing=1.6,
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.8))

    ax.set_xlim(xlim)
    ax.set_ylim(ylim)
    ax.set_xlabel(r"ln|j| (j in A/cm$^2$)", fontsize=10)
    ax.set_ylabel(r"Total Overpotential ($\eta$)", fontsize=10)
    ax.set_title(title, fontsize=9, pad=6)
    ax.xaxis.set_major_locator(plt.MultipleLocator(1))
    ax.yaxis.set_major_locator(plt.MultipleLocator(0.1))
    ax.tick_params(which="both", direction="in", top=True, right=True)
    ax.minorticks_on()

    plt.tight_layout()
    out = os.path.join(OUT, outname)
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {out}\n")


os.makedirs(OUT, exist_ok=True)

plot_tafel(
    cond_folder = "O2 BP",
    dur_dir     = "BOL",
    fname       = "CN BM 1.5bp O2",
    title       = "CN BM — O$_2$, 1.5 bar$_g$ — BOL",
    outname     = "Tafel_CNBM_O2BP_BOL.png",
    xlim        = (-5, 1),
    ylim        = (-1.05, -0.35),
)
