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
OUT  = r"Overpotnital\O2 BP\BOL\Edited\CN BM 1.5bp O2_overpot_breakdown.png"

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
    matplotlib.rcParams.update({
        "font.family": "sans-serif",
        "font.size": 10,
        "axes.linewidth": 1.0,
    })

    fig, ax = plt.subplots(figsize=(5.5, 4.5))

    # Colors matching reference image
    c_kin  = "#FFA07A"   # salmon / orange
    c_ohm  = "#90C090"   # medium green
    c_pro  = "#9090D8"   # blue-purple
    c_mass = "#F0E860"   # yellow

    # Fill regions (stacked)
    ax.fill_between(I_mA, V_kin,  V_erev,   color=c_kin,  alpha=0.85, label="Kinetic")
    ax.fill_between(I_mA, V_ohm,  V_kin,    color=c_ohm,  alpha=0.85, label="Ohmic")
    ax.fill_between(I_mA, V_pro,  V_ohm,    color=c_pro,  alpha=0.85, label="Proton")
    ax.fill_between(I_mA, V_meas, V_pro,    color=c_mass, alpha=0.85, label="Mass Transport")

    # Measured I-V curve
    ax.plot(I_mA, V, color="black", linewidth=1.5, zorder=5)

    # Erev reference line (dashed, light grey)
    ax.axhline(erev, color="grey", linewidth=0.8, linestyle="--", alpha=0.6)

    # Axis limits
    ax.set_xlim(0, 3500)
    ax.set_ylim(0.2, 1.2)

    ax.set_xlabel(r"Current Density (mAcm$^{-2}$)", fontsize=10)
    ax.set_ylabel("Cell Voltage (V)", fontsize=10)

    # Legend (top right, no frame)
    ax.legend(loc="upper right", fontsize=8, frameon=True,
              framealpha=0.9, edgecolor="grey")

    ax.tick_params(direction="in", top=True, right=True)

    plt.tight_layout()
    plt.savefig(out, dpi=200, bbox_inches="tight")
    print(f"Saved → {out}")
    plt.show()


if __name__ == "__main__":
    main()
