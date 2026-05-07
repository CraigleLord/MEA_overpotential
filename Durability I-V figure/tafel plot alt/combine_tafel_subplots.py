import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from pathlib import Path

# Panels rendered by render_tafel_panels.py (large fonts, no embedded labels)
folder = Path(__file__).parent / "Tafel_large"

plot1_files = [
    "Tafel_CNBM_O2BP_alt.png",
    "Tafel_KBBM_O2BP_alt.png",
    "Tafel_VCPolyol_O2BP_alt.png",
    "Tafel_ABPolyol_O2BP_alt.png",
]

plot2_files = [
    "Tafel_CNPolyol_O2BP_alt.png",
    "Tafel_KBPolyol_O2BP_alt.png",
    "Tafel_VCBM_O2BP_alt.png",
    "Tafel_ABBM_O2BP_alt.png",
]

XLABEL = r"ln|j|  (j in A cm$^{-2}$)"
YLABEL = r"Total Overpotential ($\eta$)"


def make_combined(files, out_name):
    imgs = [mpimg.imread(folder / f)[..., :3] for f in files]

    fig, axes = plt.subplots(1, 4, figsize=(20, 5))

    for ax, img in zip(axes, imgs):
        ax.imshow(img, aspect="auto")
        ax.axis("off")

    plt.tight_layout(rect=[0.09, 0.10, 1, 1])

    # Shared x-axis label centered across all panels
    xc = (axes[0].get_position().x0 + axes[-1].get_position().x1) / 2
    yb = axes[0].get_position().y0 - 0.02
    fig.text(xc, yb, XLABEL, ha="center", va="top", fontsize=26)

    # Shared y-axis label on the left
    yc = (axes[0].get_position().y0 + axes[0].get_position().y1) / 2
    xl = axes[0].get_position().x0 - 0.012
    fig.text(xl, yc, YLABEL, ha="center", va="center", fontsize=26, rotation=90)

    out_path = Path(__file__).parent / out_name
    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out_path}")


make_combined(plot1_files, "Tafel_combined_BM_Polyol_mix.png")
make_combined(plot2_files, "Tafel_combined_Polyol_BM_rest.png")
