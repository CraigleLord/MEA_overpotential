"""Run only build_grid for the two KB figures, saving to Final Plots/."""
import os

HERE = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(HERE, "plot_lsv_kb_dur_grid.py"), encoding="utf-8") as fh:
    code = fh.read()

# Execute everything up to (but not including) the top-level build_* calls.
cutoff = code.index('\nbuild_grid("KB BM"')
ns = {"__name__": "__main__", "__file__": os.path.join(HERE, "plot_lsv_kb_dur_grid.py")}
exec(compile(code[:cutoff], "plot_lsv_kb_dur_grid.py", "exec"), ns)

FINAL = os.path.join(ns["OUT_DIR"], "Final Plots")
os.makedirs(FINAL, exist_ok=True)

ns["build_grid"]("KB BM",     ns["MAIN_ORDER"],  ns["MAIN_CLR"],  "LSV_KB_BM_dur_grid.png",     out_dir=FINAL)
ns["build_grid"]("KB Polyol", ns["OTHER_ORDER"], ns["OTHER_CLR"], "LSV_KB_Polyol_dur_grid.png", out_dir=FINAL)
print("Done.")
