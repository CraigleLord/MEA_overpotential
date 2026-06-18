# Session handoff — overpotential SI bar charts

**Recall keyword: `overpot-barcharts`**  (also the git branch name and the tag in this commit message)

On another machine: `git fetch origin && git checkout overpot-barcharts`, then tell Claude
"continue the **overpot-barcharts** work — read HANDOFF_overpot-barcharts.md".

Date: 2026-06-18. Repo: CraigleLord/MEA_overpotential. Project root: `For paper SI/`.
All figures output to `260603 New data set (reproducibility)/Final Plots/`.

---

## What this session produced

All bar-chart work for the PEMFC catalyst SI (CN/KB/VC/AB carbons, BM & Polyol routes,
Air/O₂ × 0/1.5 bar BP, BOL vs 75K AST).

### Main workhorse script: `plot_overpot_mt_bar_grid.py`
Builds nearly everything below. Key internals:
- **Caches** (built once at top):
  - `cache` / `cache75` — BOL / 75K I-V + params, using `find_file` which routes
    **VC Polyol → `Overpotential DATA raw final/`** (final dataset) and the 260603
    repro overrides (KB BM Air BP, O₂ BP CN/KB/VC/AB). Used by the **component** charts.
  - `cache_g` / `cache_g75` — **grid-resolution** caches via `find_file_grid`
    (only the KB BM Air BP override; **VC Polyol from ALT_DATA**, matching the Overpot
    grids). Used by the **physical j-V / V-j waterfall** charts so their numbers equal
    the red marker values in `Overpot_grid_*`.
- `calc_components_at_voltage` / `calc_components_at_current` — 4 overpotential
  components (kinetic, ohmic, proton, mass-transport) at a target V or current.
- Component colours: Kinetic `#FFA500`, Ohmic `#90C090`, Proton `#9090D8`, MT `#F0E860`.
- 75K darker shade for waterfall = `#555555`, BOL light = `#cccccc`.

### Figure families (all in `Final Plots/`)
- **Component breakdown** (`build_combined_component_grid`): 4 cond rows × 2 cols,
  per-sample bars of the 4 components. Variants:
  - `Bar_Components_{Main,Other}.png` — 8 bars/sample (BOL solid + 75K hatched), V cols (0.7/0.4 V).
  - `Bar_Components_0p3_*` — same, y 0–0.3 zoom.
  - `Bar_Components_J_{Main,Other}.png` — 8 bars, current cols (0.2/1.0 A cm⁻²).
  - `Bar_Components_J_75K_Main.png` — 75K only, 4 bars, current cols, no legend.
  - `Bar_Components_75K_Main.png` — 75K only, 4 bars, V cols.
  - VC Polyol here uses the **final** dataset.
- **Waterfall j-V / V-j** (`build_waterfall_grid`, 4 own folders): 4 cond rows × 2 cols,
  per-sample BOL/75K. `style=` waterfall / grouped / bol / k75.
  - Folders: `Waterfall {Main,Other} {current density,cell voltage}/`
  - `Bar_Waterfall_*` (waterfall: 75K dark bottom + BOL increment; **BOL label on top,
    75K inside** — switched last), `Bar_Grouped_*` (side-by-side), `Bar_BOL_*`, `Bar_75K_*`.
  - These use **grid-resolution** data (`cache_g`).
- `build_stacked_dur_grid` → `Bar_Stack_BOL75K_{0p7V,0p4V}_{Main,Other}.png`
  (stacked components, BOL vs 75K grouped, per condition row).
- `build_75k_jV_grid` → `Main_barchart/Bar_75K_jV_Main.png` (75K-only single bars).
- Overpot decomposition grids: `Overpot_grid_{Main,Other}_300mA.png` (BOL),
  `Overpot_grid_{Main,Other}_75K.png` (75K) — from `plot_overpot_grid_300mA_marker.py`.
  Red markers at 0.7 V / 0.4 V labelled with current density (2 d.p.).

### Other scripts touched
- `plot_bar_individual_panels.py` — `Main_barchart/` and `Other 0.7_0.4V bar chart/`
  individual BOL bar panels. KB BM Air BP routed to 260603 repro alt (gives 0.40 A cm⁻²
  at 0.7 V, was 0.47). y: 0.7 V→0–0.8/0.2; 0.4 V→0–2.5/0.5.
- `plot_overpot_dur_300mA_marker.py` — `Bar_Eta_{0p5,1p0}_*` component columns
  (0p5 = at 0.7 V, 1p0 = at 0.4 V). Row-1 (Kinetic) y-axis fixed 0.40–0.60, 0.05/tick.

## Open decision (flagged to user, unresolved)
VC Polyol is intentionally sourced **two ways**: the **final** dataset for the component/
overpotential charts, but **ALT_DATA** (grid resolution) for the physical j-V/V-j waterfall
charts (so they match `Overpot_grid_*` red values). If everything should use one VC Polyol
dataset, align `find_file_grid` (and the grids) accordingly.
