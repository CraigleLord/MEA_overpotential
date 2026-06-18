"""Add publication panel letters (a), b), …) to each figure in Final Plots/.

Letters reset within each figure-type group and follow CN→KB→VC→AB sample order.
Labeled copies are saved to Final Plots/Labeled/.
"""

import os
from PIL import Image, ImageDraw, ImageFont

FOLDER = os.path.dirname(os.path.abspath(__file__))
OUT_FOLDER = os.path.join(FOLDER, "Labeled")
os.makedirs(OUT_FOLDER, exist_ok=True)

# Each tuple: (filename, letter)
# Letters reset per group; within group ordered CN→KB→VC→AB, BM before Polyol.
GROUPS = [
    ("IV_PD", [
        ("IV_PD_CN_BM_grid.png",     "a"),
        ("IV_PD_CN_Polyol_grid.png", "b"),
        ("IV_PD_KB_BM_grid.png",     "c"),
        ("IV_PD_KB_Polyol_grid.png", "d"),
        ("IV_PD_VC_BM_grid.png",     "e"),
        ("IV_PD_VC_Polyol_grid.png", "f"),
        ("IV_PD_AB_BM_grid.png",     "g"),
        ("IV_PD_AB_Polyol_grid.png", "h"),
    ]),
    ("LSV_dur", [
        ("LSV_KB_BM_dur_grid.png",     "a"),
        ("LSV_KB_Polyol_dur_grid.png", "b"),
    ]),
    ("LSV_combined", [
        ("LSV_combined_2x4_Main_0p7V.png",  "a"),
        ("LSV_combined_2x4_Main_0p4V.png",  "b"),
        ("LSV_combined_2x4_Other_0p7V.png", "c"),
        ("LSV_combined_2x4_Other_0p4V.png", "d"),
    ]),
    ("Overpot_dur", [
        ("Overpot_dur_Air_15bp_Main_300mA.png",  "a"),
        ("Overpot_dur_Air_15bp_Other_300mA.png", "b"),
    ]),
    ("Overpot_grid", [
        ("Overpot_grid_Main_300mA.png",  "a"),
        ("Overpot_grid_Other_300mA.png", "b"),
    ]),
    ("PD", [
        ("PD_KB_BM_grid.png", "a"),
    ]),
    ("Bar", [
        ("Bar_Other_0p7_0p4V_grid.png", "a"),
    ]),
    ("Table", [
        ("Table_O2_15bp_Main_alt.png",  "a"),
        ("Table_O2_15bp_Other_alt.png", "b"),
    ]),
    ("Tafel", [
        ("Tafel_combined_BM_Polyol_mix.png",   "a"),
        ("Tafel_combined_Polyol_BM_rest.png",  "b"),
    ]),
]

# Font: Arial Bold (Windows), fallback to DejaVu Sans Bold
FONT_PATHS = [
    r"C:\Windows\Fonts\arialbd.ttf",
    r"C:\Windows\Fonts\Arial Bold.ttf",
]

def get_font(size):
    for fp in FONT_PATHS:
        if os.path.exists(fp):
            return ImageFont.truetype(fp, size)
    return ImageFont.load_default()

def add_label(img_path, letter, out_path):
    img = Image.open(img_path).convert("RGBA")
    w, h = img.size

    # Font size: ~7% of image height, capped for very tall images
    font_size = max(60, int(min(w, h) * 0.07))
    font = get_font(font_size)

    pad = int(min(w, h) * 0.012)   # ~1.2% margin from corner
    label = f"{letter})"

    draw = ImageDraw.Draw(img)

    # Measure bounding box for the label
    bbox = draw.textbbox((0, 0), label, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]

    x, y = pad, pad

    # White halo for contrast (draw 4 times offset by 2px, then black on top)
    halo = 3
    for dx in (-halo, 0, halo):
        for dy in (-halo, 0, halo):
            if dx != 0 or dy != 0:
                draw.text((x + dx, y + dy), label, font=font, fill=(255, 255, 255, 255))
    draw.text((x, y), label, font=font, fill=(0, 0, 0, 255))

    # Save as PNG (convert back to RGB if source has no alpha)
    img.convert("RGB").save(out_path, dpi=(300, 300))
    print(f"  {os.path.basename(out_path)}  [{label}]  font_size={font_size}")

print(f"Output folder: {OUT_FOLDER}\n")
for group_name, entries in GROUPS:
    print(f"[{group_name}]")
    for fname, letter in entries:
        src = os.path.join(FOLDER, fname)
        if not os.path.exists(src):
            print(f"  MISSING: {fname}")
            continue
        dst = os.path.join(OUT_FOLDER, fname)
        add_label(src, letter, dst)
    print()

print("Done.")
