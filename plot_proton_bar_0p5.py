"""
Standalone proton-transfer overpotential bar charts at 0.5 A cm⁻²
4 rows (samples) x 1 col, for Main and Other groups.
Output: 260603 New data set (reproducibility)/Final Plots/
"""

import os, sys, math
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE     = os.path.dirname(os.path.abspath(__file__))
ALT_DATA = os.path.join(BASE, "Durability I-V figure", "tafel plot alt",
                         "Overpotential DATA raw alt")
REPRO_DATA_260603 = os.path.join(BASE, "260603 New data set (reproducibility)",
                                  "Overpotential DATA raw alt")
FINAL_PLOTS_DIR = os.path.join(BASE, "260603 New data set (reproducibility)", "Final Plots")
os.makedirs(FINAL_PLOTS_DIR, exist_ok=True)

TARGET_J = 500.0   # mA/cm² (= 0.5 A cm⁻²)

MAIN_SAMPLES  = ["CN BM",     "KB BM",     "VC Polyol", "AB Polyol"]
OTHER_SAMPLES = ["CN Polyol", "KB Polyol", "VC BM",     "AB BM"]

BOL_CONDITIONS = [
    ("Air",    "BOL"),
    ("Air BP", "BOL"),
    ("O2",     "BOL"),
    ("O2 BP",  "BOL"),
]

COND_LABELS = ["Air\n0 BP", "Air\n1.5 BP", r"O$_2$" + "\n0 BP", r"O$_2$" + "\n1.5 BP"]

FILE_STEMS = {
    ("Air",    "CN BM"):     "CN BM air",
    ("Air",    "KB BM"):     "KB BM air",
    ("Air",    "VC Polyol"): "VC Polyol air",
    ("Air",    "AB Polyol"): "AB Polyol air",
    ("Air",    "CN Polyol"): "CN Polyol air",
    ("Air",    "KB Polyol"): "KB Polyol air",
    ("Air",    "VC BM"):     "VC BM air",
    ("Air",    "AB BM"):     "AB BM Air",

    ("Air BP", "CN BM"):     "CN BM 1.5bp air",
    ("Air BP", "KB BM"):     "KB BM 1.5bp air",
    ("Air BP", "VC Polyol"): "VC Polyol 1.5bp air",
    ("Air BP", "AB Polyol"): "AB Polyol 1.5bp air",
    ("Air BP", "CN Polyol"): "CN Polyol 1.5bp air",
    ("Air BP", "KB Polyol"): "KB Polyol 1.5bp air",
    ("Air BP", "VC BM"):     "VC BM 1.5bp air",
    ("Air BP", "AB BM"):     "AB BM 1.5bp air",

    ("O2",    "CN BM"):      "CN BM o2",
    ("O2",    "KB BM"):      "KB BM O2",
    ("O2",    "VC Polyol"):  "VC Polyol O2",
    ("O2",    "AB Polyol"):  "AB Polyol O2",
    ("O2",    "CN Polyol"):  "CN Polyol O2",
    ("O2",    "KB Polyol"):  "KB Polyol O2",
    ("O2",    "VC BM"):      "VC BM O2",
    ("O2",    "AB BM"):      "AB BM  O2",

    ("O2 BP", "CN BM"):      "CN BM 1.5bp O2",
    ("O2 BP", "KB BM"):      "KB BM 1.5bp O2",
    ("O2 BP", "VC Polyol"):  "VC Polyol 1.5bp O2",
    ("O2 BP", "AB Polyol"):  "AB Polyol 1.5bp O2",
    ("O2 BP", "CN Polyol"):  "CN Polyol 1.5bp O2",
    ("O2 BP", "KB Polyol"):  "KB Polyol 1.5bp O2",
    ("O2 BP", "VC BM"):      "VC BM 1.5bp O2",
    ("O2 BP", "AB BM"):      "AB BM 1.5bp O2",
}

# EIS-derived proton resistance override (same as main script)
EIS_RPRO = {
    "CN BM": 0.020229,
}

DATA_OVERRIDES = {
    ("Air BP", "KB BM 1.5bp air"): os.path.join(REPRO_DATA_260603, "Air BP", "KB BM"),
}


def find_file(cond_folder, dur_folder, stem):
    override = DATA_OVERRIDES.get((cond_folder, stem))
    if override:
        for dur_try in (dur_folder, dur_folder.lower()):
            for sfx in ("_edited.xlsx", "_unchanged.xlsx"):
                p = os.path.join(override, dur_try, stem + sfx)
                if os.path.exists(p):
                    return p
    for dur_try in (dur_folder, dur_folder.lower()):
        folder = os.path.join(ALT_DATA, cond_folder, dur_try, "Edited")
        for sfx in ("_edited.xlsx", "_unchanged.xlsx"):
            p = os.path.join(folder, stem + sfx)
            if os.path.exists(p):
                return p
    return None


def read_pro(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=False)
    ws = wb["Sheet1"]
    rows = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        a, e = row[0], row[4]
        if isinstance(a, (int, float)) and a > 0:
            rows.append((a, float(e) if isinstance(e, (int, float)) else np.nan))
    wb.close()
    if len(rows) < 5:
        raise ValueError("Too few data rows")
    I_mA  = np.array([r[0] * 200 for r in rows])
    R_pro = np.array([r[1]        for r in rows])
    return I_mA, R_pro


def get_eta_pro_at_target(sample, cond_folder, dur_folder):
    stem = FILE_STEMS.get((cond_folder, sample))
    fp   = find_file(cond_folder, dur_folder, stem) if stem else None
    if fp is None:
        print(f"  WARN: file not found: {cond_folder}/{dur_folder}/{stem}")
        return np.nan
    try:
        I_mA, R_pro = read_pro(fp)
        if sample in EIS_RPRO:
            R_pro = np.full(len(R_pro), EIS_RPRO[sample])
        I_A     = I_mA / 200          # same convention as draw_panel
        eta_pro = R_pro * I_A
        val = float(np.interp(TARGET_J, I_mA, eta_pro))
        print(f"  {sample}/{cond_folder}: η_pro = {val:.4f} V")
        return val
    except Exception as e:
        print(f"  ERROR {sample}/{cond_folder}: {e}")
        return np.nan


def build_bar_chart(sample_list, group_name):
    FS = 22

    plt.rcParams.update({
        "font.family":      "Arial",
        "axes.linewidth":   1.0,
        "xtick.direction":  "out",
        "ytick.direction":  "out",
        "xtick.top":        False,
        "ytick.right":      False,
    })

    # Collect all values first for shared y-scale
    all_data = []
    for samp in sample_list:
        row = [get_eta_pro_at_target(samp, cf, df) for cf, df in BOL_CONDITIONS]
        all_data.append(row)

    valid = [v for row in all_data for v in row if not np.isnan(v)]
    max_val = max(valid) if valid else 0.06

    # Compute a "round" step so that exactly 3 intervals (4 ticks) span the data
    raw_step = max_val / 3
    exp = math.floor(math.log10(raw_step))
    base = 10 ** exp
    nice_step = next(m * base for m in [1, 2, 2.5, 5, 10] if m * base >= raw_step)
    ymax = nice_step * 3 + nice_step * 0.15   # small buffer above top tick

    n_rows = len(sample_list)
    n_cond = len(BOL_CONDITIONS)
    x      = np.arange(n_cond)

    fig = plt.figure(figsize=(5.5, 3.5 * n_rows), constrained_layout=True)
    gs  = fig.add_gridspec(n_rows, 1)

    for row_idx, (samp, vals) in enumerate(zip(sample_list, all_data)):
        ax = fig.add_subplot(gs[row_idx, 0])

        ax.bar(x, vals, width=0.55,
               facecolor="#cccccc", edgecolor="black", linewidth=0.5)

        ax.set_ylim(0, ymax)
        ax.set_xlim(-0.5, n_cond - 0.5)
        ax.tick_params(direction="out", top=False, right=False,
                       labelsize=FS, width=1.0, length=4)
        ax.yaxis.set_major_locator(plt.MultipleLocator(nice_step))
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: f"{y:.2f}"))
        for spine in ax.spines.values():
            spine.set_linewidth(1.0)

        # x-tick labels only on last row
        ax.set_xticks(x)
        if row_idx == n_rows - 1:
            ax.set_xticklabels(COND_LABELS, fontsize=FS)
        else:
            ax.set_xticklabels([""] * n_cond)

        # No title, no axis labels
        ax.set_title("")
        ax.set_xlabel("")
        ax.set_ylabel("")

    fname = f"Proton_overpot_bar_{group_name}_0p5Acm2.png"
    out   = os.path.join(FINAL_PLOTS_DIR, fname)
    fig.savefig(out, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {out}")


build_bar_chart(MAIN_SAMPLES,  "Main")
build_bar_chart(OTHER_SAMPLES, "Other")
print("Done.")
