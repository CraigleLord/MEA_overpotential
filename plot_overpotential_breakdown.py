"""
Overpotential breakdown plot for fuel cell polarization curves.
Stacked-region style (kinetic / ohmic / proton / mass-transport).

Currently configured for: CN BM, O2, 1.5 bar BP, BOL
"""

import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import openpyxl
import math
import os

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

FILE = (
    r"Overpotnital\O2 BP\BOL\Edited\CN BM 1.5bp O2_edited.xlsx"
)
OUT  = r"Overpotential Graphs\CN BM 1.5bp O2_overpot_breakdown.png"

# Tafel parameters from fitted image (BOL)
TAFEL_INTERCEPT = -0.517   # V  (η = slope·ln j + intercept, j in A cm⁻²)
TAFEL_SLOPE     = -0.0358  # V per ln-unit

# Cell conditions (same as in the spreadsheet)
T_K    = 353    # K
P_H2   = 250    # kPa
P_O2   = 250    # kPa

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def calc_erev(T, p_h2_kPa, p_o2_kPa):
    """Nernst-corrected reversible potential (same formula as the Excel)."""
    p_h2_rel = p_h2_kPa / 101.3
    p_o2_rel = p_o2_kPa / 101.3
    # S = ((p_H2)^2 * p_O2)^2  (activity product for the Nernst term)
    S = ((p_h2_rel ** 2) * p_o2_rel) ** 2
    erev = 1.23 - 0.0009 * (T - 298) + (2.303 * 8.314 * T / (4 * 96485)) * math.log10(S)
    return erev


def read_data(filepath):
    """
    Read Sheet1 and return arrays:
      I_mA  : current density  [mA cm⁻²]
      V     : cell voltage     [V]
      R_ohm : ohmic resistance [Ω] (total, not ASR)
      R_pro : proton resistance [Ω]
    """
    wb = openpyxl.load_workbook(filepath, data_only=False)
    ws = wb["Sheet1"]

    I_list, V_list, R_ohm_list, R_pro_list = [], [], [], []

    for row in ws.iter_rows(min_row=4, values_only=True):
        i_raw = row[0]   # Col A: total current [A]
        v_raw = row[1]   # Col B: cell voltage [V]
        r_ohm = row[3]   # Col D: ohmic resistance [Ω]
        r_pro = row[4]   # Col E: proton resistance [Ω]

        if not isinstance(i_raw, (int, float)):
            continue
        if i_raw <= 0 or v_raw is None:
            continue

        # F = A * 200  → I_density in mA cm⁻²  (area = 5 cm²)
        i_density = i_raw * 200   # mA cm⁻²
        I_list.append(i_density)
        V_list.append(float(v_raw))
        R_ohm_list.append(float(r_ohm) if r_ohm is not None else np.nan)
        R_pro_list.append(float(r_pro) if r_pro is not None else np.nan)

    return (
        np.array(I_list),
        np.array(V_list),
        np.array(R_ohm_list),
        np.array(R_pro_list),
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    here = os.path.dirname(os.path.abspath(__file__))
    fp   = os.path.join(here, FILE)
    out  = os.path.join(here, OUT)

    I_mA, V, R_ohm, R_pro = read_data(fp)
    I_A_total = I_mA / 200      # total current [A]  (reverse of A*200)
    I_Acm2    = I_mA / 1000     # current density [A cm⁻²]

    erev = calc_erev(T_K, P_H2, P_O2)
    print(f"Erev = {erev:.4f} V")

    # --- Overpotential components ---
    ln_j = np.log(I_Acm2)   # ln(j) with j in A cm⁻²

    eta_kin  = -(TAFEL_INTERCEPT + TAFEL_SLOPE * ln_j)  # kinetic  [V, positive loss]
    eta_ohm  = R_ohm * I_A_total                          # ohmic    [V, positive loss]
    eta_pro  = R_pro * I_A_total                          # proton   [V, positive loss]
    eta_mass = (erev - V) - eta_kin - eta_ohm - eta_pro  # residual [V, positive loss]

    # Stacked voltage levels (top → bottom)
    V_erev    = np.full_like(I_mA, erev)
    V_kin     = erev - eta_kin               # after kinetic loss
    V_ohm     = V_kin  - eta_ohm            # after + ohmic
    V_pro     = V_ohm  - eta_pro            # after + proton
    V_meas    = V                             # measured (= after + mass-transport)

    # --- Plot ---
    # Figure is ~5.5 x 4.5" so that proportions match the reference image.
    # FS is set relative to figure inches so fonts are visually large.
    FS   = 14    # pt — applied to every text element individually
    LW   = 1.5   # I-V curve line width
    BW   = 1.0   # axis border + tick line width
    DPI  = 300

    fig, ax = plt.subplots(figsize=(5.5, 4.5))

    # Reset rcParams cleanly so nothing bleeds in from previous state
    matplotlib.rcParams.update(matplotlib.rcParamsDefault)
    matplotlib.rcParams.update({
        "font.family":     "sans-serif",
        "font.size":       FS,
        "axes.labelsize":  FS,
        "axes.titlesize":  FS,
        "xtick.labelsize": FS,
        "ytick.labelsize": FS,
        "legend.fontsize": FS,
        "axes.linewidth":    BW,
        "xtick.major.width": BW,
        "ytick.major.width": BW,
    })

    # Colors matching reference image
    c_kin  = "#FFA07A"   # salmon / orange
    c_ohm  = "#90C090"   # medium green
    c_pro  = "#9090D8"   # blue-purple
    c_mass = "#F0E860"   # yellow

    # Clip mass-transport: no yellow below the I-V curve
    V_pro_clipped = np.maximum(V_pro, V_meas)

    # Filled regions (stacked)
    ax.fill_between(I_mA, V_kin,  V_erev,        color=c_kin,  alpha=0.85, label="Kinetic")
    ax.fill_between(I_mA, V_ohm,  V_kin,         color=c_ohm,  alpha=0.85, label="Ohmic")
    ax.fill_between(I_mA, V_pro,  V_ohm,         color=c_pro,  alpha=0.85, label="Proton")
    ax.fill_between(I_mA, V_meas, V_pro_clipped, color=c_mass, alpha=0.85, label="Mass Transport")

    # Measured I-V curve
    ax.plot(I_mA, V, color="black", linewidth=LW, zorder=5)

    # --- Red dots at 500 mA cm⁻² ---
    i_mark      = 500.0
    v_kin_mark  = float(np.interp(i_mark, I_mA, V_kin))
    v_meas_mark = float(np.interp(i_mark, I_mA, V_meas))

    dot_kw = dict(color="red", marker="o", markersize=8,
                  linestyle="none", zorder=10)
    ax.plot(i_mark, v_kin_mark,  **dot_kw)
    ax.plot(i_mark, v_meas_mark, **dot_kw)

    lbl_kw = dict(color="red", fontsize=FS, va="center", ha="left")
    ax.text(i_mark + 50, v_kin_mark,  f"{v_kin_mark:.2f}",  **lbl_kw)
    ax.text(i_mark + 50, v_meas_mark, f"{v_meas_mark:.2f}", **lbl_kw)

    # Erev reference line (dashed, light grey)
    ax.axhline(erev, color="grey", linewidth=0.8, linestyle="--", alpha=0.6)

    # Axis limits and labels
    ax.set_xlim(0, 3500)
    ax.set_ylim(0.2, 1.2)
    ax.set_xlabel(r"Current Density (mAcm$^{-2}$)", fontsize=FS)
    ax.set_ylabel("Cell Voltage (V)", fontsize=FS)

    # Tick style
    # Ticks outside, only on left and bottom axes
    ax.tick_params(direction="out", top=False, right=False,
                   labelsize=FS, width=BW, length=5)

    # Legend — white box, black border
    leg = ax.legend(loc="upper right", fontsize=FS, frameon=True,
                    facecolor="white", edgecolor="black")
    leg.get_frame().set_linewidth(1.5)

    plt.tight_layout()
    plt.savefig(out, dpi=DPI, bbox_inches="tight")
    print(f"Saved → {out}")
    plt.show()


if __name__ == "__main__":
    main()
