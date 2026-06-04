# Project: Overpotential SI figures — KAIST PEMFC catalyst study

Comparing CN/KB/VC/AB carbon supports (BM and Polyol synthesis routes)
across Air/O₂ × 0/1.5 bar back-pressure at BOL/30K/75K AST cycling.

## General graphing skill

Use the conventions below whenever drawing any publication figure. They were
validated iteratively and Dave does not want to re-explain them.

### Layout
- `plt.figure(constrained_layout=True)` + `GridSpec` — never `plt.subplots()`.
- Font size: **FS = 22** for everything (ticks, labels, legend, annotations).
- Set via `plt.rcParams`: font=Arial, axes.linewidth=1.0, direction=out, top/right ticks off.

### Current density units
Always display in **A cm⁻²** (divide mA/cm² by 1000), **1 d.p.** tick labels.

### I-V subplot
- Log x-axis, `xlim=(0.005, 4)` A cm⁻²
- Formatter: `f"{x:.1f}" if x >= 0.1 else f"{x:.2f}"`
- y: linear, `ylim=(0.2, 1.0)`, `MultipleLocator(0.2)`

### Power density subplot
- Same log x-axis as I-V (`xlim=(0.005, 4)`)
- y: W cm⁻², `MaxNLocator(nbins=4, steps=[1,2,5,10])`, 1 d.p.

### Bar chart subplots
- Units: A cm⁻², `MaxNLocator(nbins=4)`, 1 d.p.
- Waterfall stack: 75K at bottom, BOL increment on top (`max(BOL−75K, 0)`)
- Bar width 0.55; x-tick sample names only on last row, split with `\n`
- Compute shared y-scale across ALL samples before any plotting

```python
BAR_STYLES = {
    "BOL": dict(facecolor="#cccccc", hatch="",  edgecolor="black", linewidth=0.5),
    "75K": dict(facecolor="#555555", hatch="",  edgecolor="black", linewidth=0.5),
}
BAR_DUR_ORDER = ["75K", "BOL"]   # bottom → top
```

### Durability line styles (30K excluded from all figures)
```python
DUR_META = {
    "BOL": dict(ls="-",  marker="s", color="#000000"),   # solid black
    "75K": dict(ls="--", marker="^", color="#444444"),   # dashed dark-gray
}
```

### Marker conventions

| Symbol | Marker code | Meaning |
|--------|-------------|---------|
| ■ | `"s"` | BOL current at voltage intersection (0.7 V / 0.4 V) |
| ▲ | `"^"` | 75K current at voltage intersection |
| ◆ | `"D"` | BOL max power density point |
| ✚ | `"P"` | 75K max power density point |

Peak/intersection markers: `ms=9`, black or white edge (`markeredgewidth=0.7/1.2`).

Voltage reference lines: `V_REF=[0.70, 0.40]`, colors `["#C62828","#1565C0"]`.
Labels go **below** the line: `va="top"`, offset `v_tgt - 0.025`, `fontsize=FS-2`.

### Legend
```python
ax.legend(loc="upper right", bbox_to_anchor=(0.76, 0.97),
          frameon=False, fontsize=FS, handlelength=1.5,
          borderpad=0.2, labelspacing=0.3)
```
Condition annotation at `(0.97, 0.97)`, `ha="right"`, `va="top"`.

### Single centred y-label per column (freeze-and-centre)
```python
# 1. set_ylabel on ALL rows so constrained_layout reserves the right space
for ax in col_axes:
    ax.set_ylabel("Label", fontsize=FS)

# 2. draw → freeze layout
fig.canvas.draw()
renderer = fig.canvas.get_renderer()
fig.set_layout_engine("none")

# 3. compute centre and place one fig.text; hide per-row labels
y_c    = (col_axes[0].get_position().y1 + col_axes[-1].get_position().y0) / 2
lbl_bb = col_axes[0].yaxis.label.get_window_extent(renderer)
x_c    = (lbl_bb.x0 + lbl_bb.x1) / 2 / (fig.get_figwidth() * fig.dpi)
for ax in col_axes:
    ax.yaxis.label.set_visible(False)
fig.text(x_c, y_c, "Label", ha="center", va="center", rotation=90, fontsize=FS)
```

## Canonical scripts

| Script | Purpose |
|--------|---------|
| `plot_lsv_kb_dur_grid.py` | All matrix grid figures (I-V, PD, bar charts) |
| `plot_overpot_dur_300mA_marker.py` | Overpotential durability grid |
| `plot_overpot_grid_300mA_marker.py` | Overpotential BOL grid |
| `plot_lsv_combined_2x4.py` | LSV combined 2×4 with bar charts |

## 260603 reproducibility data

New VC Polyol (EG) MEA data lives in `260603 New data set (reproducibility)/`.
Route VC Polyol reads to the staging files via `override_vc=True` in `load_iv()`.
All new figures output to that folder; originals in `Figure Plots (Claude AI)/` untouched.

## Data pipeline

- Raw I-V xlsx: `Overpotnital/{cond}/{dur}/Edited/{stem}_{edited|unchanged}.xlsx`
- Col A = current (Amps), Col B = voltage (V), data from row 4
- `load_iv` returns `I_d` in mA/cm²; divide by 1000 for A/cm² display
- Overpotential xlsx (Sheet1): S4=intercept, S5=slope, R2=T_K, T2=P_H2, U2=P_O2

## Figure types

| Function | Output name pattern | Description |
|----------|---------------------|-------------|
| `build_grid` | `LSV_*_dur_grid.png` | 4×3: I-V + j@0.7V bar + j@0.4V bar |
| `build_pd_grid` | `PD_*_grid.png` | 4×1: power density |
| `build_iv_pd_grid` | `IV_PD_*_grid.png` | 4×2: I-V + power density |
| `build_bar_only_grid` | `Bar_*_grid.png` | 4×2: j@0.7V bar + j@0.4V bar |
