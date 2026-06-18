"""
Illustrative CO-stripping voltammogram for AB EG, AST (after AST).

Curve shape matches the AB EG BOL reference (rise from 0, sharp negative
transient spike near 0.04-0.05 V, flat plateau, CO-oxidation peak near
0.6 V, flat plateau to 1.2 V). For AST, the two "peak" features (the
negative transient spike and the CO-oxidation peak) are dampened to 20% of
their BOL amplitude (80% reduction); the rest of the curve shape is
unchanged.

If used in the SI, caption it as a representative/schematic voltammogram,
not raw acquired data.
"""

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator

plt.rcParams.update({
    "font.family": "Arial",
    "font.size": 14,
    "axes.linewidth": 1.2,
    "xtick.direction": "in",
    "ytick.direction": "in",
    "xtick.top": True,
    "ytick.right": True,
})

x = np.linspace(0.0, 1.2, 800)

dampen = 0.20  # 80% reduction -> 20% remaining

# --- shared shape: rise from 0 to a flat plateau, plateau drifts slightly up ---
def plateau(v):
    return 0.020 + 0.004 * v


def rise_frac(v):
    return 1.0 / (1.0 + np.exp(-(v - 0.06) / 0.012))


def baseline_curve(v):
    return plateau(v) * rise_frac(v)


# --- "peak" features ---
def neg_spike(v, depth):
    return -depth * np.exp(-((v - 0.045) / 0.012) ** 2)


def co_peak(v, amp):
    return amp * np.exp(-((v - 0.60) / 0.055) ** 2)


depth_BOL = 0.0545   # negative transient spike reaches ~-0.05
amp_BOL = 0.0646     # CO peak reaches ~0.087

depth_AST = dampen * depth_BOL
amp_AST = dampen * amp_BOL

# CO-covered (1st) scan: rise + spike + CO peak
curve_CO_AST = baseline_curve(x) + neg_spike(x, depth_AST) + co_peak(x, amp_AST)

# CO-free (2nd) baseline scan: just the rise + plateau, no peaks
curve_base_AST = baseline_curve(x)

# light measurement-style noise
rng = np.random.default_rng(7)
curve_CO_AST_n = curve_CO_AST + rng.normal(0, 0.0006, x.size)
curve_base_AST_n = curve_base_AST + rng.normal(0, 0.0006, x.size)

fig, ax = plt.subplots(figsize=(4.4, 3.6), constrained_layout=True)
ax.plot(x, curve_CO_AST_n, color="black", lw=1.1)
ax.plot(x, curve_base_AST_n, color="black", lw=1.1)

ax.set_xlim(0.0, 1.2)
ax.set_ylim(-0.05, 0.20)
ax.xaxis.set_major_locator(MultipleLocator(0.2))
ax.yaxis.set_major_locator(MultipleLocator(0.05))
ax.set_xlabel("Cell Voltage", fontsize=14, fontweight="bold")
ax.set_ylabel("Current (A)", fontsize=14, fontweight="bold")
ax.tick_params(labelsize=13)

out_path = "Other claude task/AB_EG_AST_CO_stripping.png"
fig.savefig(out_path, dpi=300)
print(f"saved {out_path}")
print(f"neg spike depth: BOL={depth_BOL:.4f} -> AST={depth_AST:.4f}")
print(f"CO peak amp:     BOL={amp_BOL:.4f} -> AST={amp_AST:.4f}")
