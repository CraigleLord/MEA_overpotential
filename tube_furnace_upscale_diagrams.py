import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Arc, Wedge
import matplotlib.patheffects as pe
import numpy as np

fig = plt.figure(figsize=(18, 22))
fig.patch.set_facecolor('#f8f8f8')

FURNACE_RED   = '#c0392b'
FURNACE_LIGHT = '#f1948a'
BOAT_GRAY     = '#7f8c8d'
POWDER_BROWN  = '#a0522d'
POWDER_LIGHT  = '#d4a57a'
AR_BLUE       = '#2980b9'
TE_PURPLE     = '#8e44ad'
ARROW_COLOR   = '#2c3e50'
BG_WHITE      = '#ffffff'
MESH_COLOR    = '#1abc9c'

def furnace_shell(ax, x, y, w, h, label='Tube Furnace'):
    # Outer furnace body
    rect = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.05",
                           fc=FURNACE_LIGHT, ec=FURNACE_RED, lw=2.5)
    ax.add_patch(rect)
    # Hot zone indicator (darker center band)
    hot = FancyBboxPatch((x+w*0.15, y), w*0.7, h, boxstyle="round,pad=0.0",
                          fc=FURNACE_RED, ec='none', alpha=0.25)
    ax.add_patch(hot)
    ax.text(x+w/2, y+h+0.07, label, ha='center', va='bottom',
            fontsize=10, fontweight='bold', color=FURNACE_RED)
    ax.text(x+w/2, y-0.12, 'hot zone', ha='center', va='top',
            fontsize=8, color=FURNACE_RED, style='italic')

def tube(ax, x, y, w, h):
    rect = plt.Rectangle((x, y), w, h, fc=BG_WHITE, ec='#aaa', lw=1.5, zorder=2)
    ax.add_patch(rect)

def powder_boat(ax, x, y, w, depth, color=POWDER_BROWN, alpha=1.0, mesh=False):
    boat_h = depth + 0.08
    # boat walls
    boat = plt.Rectangle((x, y), w, boat_h, fc='#bdc3c7', ec=BOAT_GRAY, lw=1.2, zorder=3)
    ax.add_patch(boat)
    if mesh:
        mesh_rect = plt.Rectangle((x, y+0.04), w, 0.04,
                                   fc=MESH_COLOR, ec='#16a085', lw=1, zorder=4,
                                   hatch='///')
        ax.add_patch(mesh_rect)
        powder = plt.Rectangle((x, y+0.08), w, depth,
                                 fc=color, ec='none', alpha=alpha, zorder=4)
        ax.add_patch(powder)
    else:
        powder = plt.Rectangle((x, y+0.04), w, depth,
                                 fc=color, ec='none', alpha=alpha, zorder=4)
        ax.add_patch(powder)

def arrow_ar(ax, x1, x2, y, label='Ar', color=AR_BLUE):
    ax.annotate('', xy=(x2, y), xytext=(x1, y),
                arrowprops=dict(arrowstyle='->', color=color, lw=2.0))
    ax.text((x1+x2)/2, y+0.07, label, ha='center', va='bottom',
            fontsize=9, color=color, fontweight='bold')

def section_title(ax, x, y, title, subtitle=''):
    ax.text(x, y, title, fontsize=13, fontweight='bold', color='#2c3e50',
            va='top', ha='left')
    if subtitle:
        ax.text(x, y-0.18, subtitle, fontsize=9, color='#7f8c8d',
                va='top', ha='left', style='italic')

# ─────────────────────────────────────────────────
# PANEL 1 — Multi-boat thin layers
# ─────────────────────────────────────────────────
ax1 = fig.add_axes([0.04, 0.77, 0.92, 0.21])
ax1.set_xlim(0, 10); ax1.set_ylim(0, 2.2); ax1.axis('off')

section_title(ax1, 0.05, 2.15, '1  Multi-Boat Thin Layer Loading',
              'Keep powder bed ≤2 mm deep — spread mass across multiple boats in series')

# BAD example
ax1.text(1.5, 1.75, '✗  Thick single boat (bad)', fontsize=9,
         color='#c0392b', fontweight='bold', ha='center')
furnace_shell(ax1, 0.2, 0.35, 2.6, 0.9)
tube(ax1, 0.2, 0.45, 2.6, 0.7)
powder_boat(ax1, 0.65, 0.5, 1.7, 0.55, alpha=0.9)  # deep pile
ax1.annotate('~10 mm deep', xy=(1.5, 0.54+0.55), xytext=(1.5, 1.22),
             ha='center', va='bottom', fontsize=8, color=POWDER_BROWN,
             arrowprops=dict(arrowstyle='->', color=POWDER_BROWN, lw=1.2))
ax1.text(1.5, 0.52, 'bottom powder\nstarved of Te', ha='center', va='center',
         fontsize=7.5, color='white', fontweight='bold')
arrow_ar(ax1, 0.05, 0.65, 0.82, label='Ar+Te')

# GOOD example
ax1.text(6.5, 1.75, '✓  Thin boats in series (good)', fontsize=9,
         color='#27ae60', fontweight='bold', ha='center')
furnace_shell(ax1, 3.5, 0.35, 5.9, 0.9, label='')
ax1.text(6.45, 1.37, 'Tube Furnace', ha='center', va='bottom',
         fontsize=10, fontweight='bold', color=FURNACE_RED)
tube(ax1, 3.5, 0.45, 5.9, 0.7)

xs = [4.05, 5.05, 6.05, 7.05, 8.05]
for i, bx in enumerate(xs):
    powder_boat(ax1, bx, 0.5, 0.7, 0.12, alpha=0.85)
    ax1.text(bx+0.35, 0.48, f'#{i+1}', ha='center', va='top', fontsize=7,
             color='white', fontweight='bold')

ax1.annotate('', xy=(xs[0], 0.67), xytext=(xs[0], 0.92),
             arrowprops=dict(arrowstyle='->', color=POWDER_BROWN, lw=1.2))
ax1.text(xs[0]+0.35, 0.95, '~2 mm', ha='center', va='bottom', fontsize=8,
         color=POWDER_BROWN)
arrow_ar(ax1, 3.2, 4.05, 0.82, label='Ar+Te')
ax1.annotate('', xy=(9.5, 0.82), xytext=(8.75, 0.82),
             arrowprops=dict(arrowstyle='->', color='#7f8c8d', lw=1.5))
ax1.text(9.55, 0.82, 'exhaust', va='center', fontsize=8, color='#7f8c8d')

# Te source boat
te_boat = plt.Rectangle((3.55, 0.5), 0.4, 0.3, fc=TE_PURPLE, ec='#6c3483',
                          lw=1.2, zorder=5, alpha=0.85)
ax1.add_patch(te_boat)
ax1.text(3.75, 0.66, 'Te', ha='center', va='center', fontsize=9,
         color='white', fontweight='bold')


# ─────────────────────────────────────────────────
# PANEL 2 — Wide flat boats
# ─────────────────────────────────────────────────
ax2 = fig.add_axes([0.04, 0.535, 0.92, 0.215])
ax2.set_xlim(0, 10); ax2.set_ylim(0, 2.2); ax2.axis('off')

section_title(ax2, 0.05, 2.15, '2  Wide Flat Boats — Maximize Tube Cross-Section',
              'Use D-shaped or wide alumina trays to fill the tube width')

def draw_tube_cross(ax, cx, cy, r_out, boat_w, boat_h, label_top, label_bot,
                    pct_label, color=POWDER_BROWN):
    circle = plt.Circle((cx, cy), r_out, fc='#ecf0f1', ec='#7f8c8d', lw=2, zorder=2)
    ax.add_patch(circle)
    bx = cx - boat_w/2; by = cy - r_out*0.5
    boat = plt.Rectangle((bx, by), boat_w, boat_h,
                           fc=color, ec=BOAT_GRAY, lw=1.2, zorder=3, alpha=0.85)
    ax.add_patch(boat)
    ax.text(cx, cy+r_out+0.12, label_top, ha='center', va='bottom',
            fontsize=9, fontweight='bold')
    ax.text(cx, cy-r_out-0.12, label_bot, ha='center', va='top',
            fontsize=8, color='#7f8c8d')
    ax.text(cx, cy, pct_label, ha='center', va='center',
            fontsize=10, fontweight='bold', color='white',
            path_effects=[pe.withStroke(linewidth=2, foreground=BOAT_GRAY)])

draw_tube_cross(ax2, 2.0, 1.1, 0.75, 0.7, 0.35,
                'Standard boat', '~40% of tube area used', '40%')
draw_tube_cross(ax2, 5.0, 1.1, 0.75, 1.35, 0.38,
                'Wide alumina tray', '~75% of tube area used', '75%',
                color='#c0392b')

# D-shaped boat
ax2.text(8.1, 2.1, 'D-shaped boat\n(fits curved tube wall)', ha='center',
         va='top', fontsize=9, fontweight='bold')
circle2 = plt.Circle((8.1, 1.1), 0.75, fc='#ecf0f1', ec='#7f8c8d', lw=2, zorder=2)
ax2.add_patch(circle2)
theta = np.linspace(-np.pi*0.75, np.pi*0.75, 60)
# D-shape: flat bottom, curved sides following tube interior
xs_d = 8.1 + 0.72*np.cos(theta)
ys_d = 1.1 + 0.72*np.sin(theta)
xs_d = np.append(xs_d, [8.1+0.72*np.cos(-np.pi*0.75)])
ys_d = np.append(ys_d, [1.1+0.72*np.sin(-np.pi*0.75)])
ax2.fill(xs_d, ys_d, fc=POWDER_BROWN, ec=BOAT_GRAY, lw=1.5, zorder=3, alpha=0.85)
ax2.text(8.1, 1.1, '~85%', ha='center', va='center',
         fontsize=10, fontweight='bold', color='white',
         path_effects=[pe.withStroke(linewidth=2, foreground=BOAT_GRAY)])
ax2.text(8.1, 0.2, '~85% of tube area used', ha='center', va='top',
         fontsize=8, color='#7f8c8d')

ax2.annotate('', xy=(3.6, 1.1), xytext=(3.0, 1.1),
             arrowprops=dict(arrowstyle='->', color='#27ae60', lw=2))
ax2.annotate('', xy=(6.6, 1.1), xytext=(6.0, 1.1),
             arrowprops=dict(arrowstyle='->', color='#27ae60', lw=2))
ax2.text(3.3, 1.35, 'upgrade', ha='center', fontsize=8,
         color='#27ae60', fontweight='bold')
ax2.text(6.3, 1.35, 'best option', ha='center', fontsize=8,
         color='#27ae60', fontweight='bold')


# ─────────────────────────────────────────────────
# PANEL 3 — Mesh boat
# ─────────────────────────────────────────────────
ax3 = fig.add_axes([0.04, 0.315, 0.92, 0.200])
ax3.set_xlim(0, 10); ax3.set_ylim(0, 2.0); ax3.axis('off')

section_title(ax3, 0.05, 1.95, '3  Mesh / Screen Boat — Flow Through the Bed',
              'Coarse alumina mesh frit lets Te vapor rise through powder from below')

# Solid boat cross-section
ax3.text(2.0, 1.85, 'Solid boat — vapor flows only over top', fontsize=9,
         fontweight='bold', ha='center', color='#c0392b')
tube(ax3, 0.3, 0.3, 3.4, 1.2)
bx, by, bw, bh = 0.8, 0.35, 2.4, 0.6
boat_s = plt.Rectangle((bx, by), bw, bh, fc=POWDER_BROWN, ec=BOAT_GRAY,
                         lw=1.5, zorder=3, alpha=0.85)
ax3.add_patch(boat_s)
solid_base = plt.Rectangle((bx, by), bw, 0.06, fc=BOAT_GRAY, ec='none', zorder=4)
ax3.add_patch(solid_base)
for xi in np.linspace(0.95, 2.95, 8):
    ax3.annotate('', xy=(xi, 0.97), xytext=(xi, 1.12),
                arrowprops=dict(arrowstyle='->', color=AR_BLUE, lw=1.0))
ax3.text(2.0, 1.18, 'vapor flows over top only', ha='center', va='bottom',
         fontsize=8, color=AR_BLUE)
ax3.text(2.0, 0.62, 'powder\n(bottom layer\nunder-exposed)', ha='center',
         va='center', fontsize=8, color='white', fontweight='bold')
arrow_ar(ax3, 0.1, 0.8, 0.75, label='Ar+Te')

# Mesh boat cross-section
ax3.text(7.5, 1.85, 'Mesh boat — vapor flows through bed', fontsize=9,
         fontweight='bold', ha='center', color='#27ae60')
tube(ax3, 4.8, 0.3, 5.1, 1.2)
bx2, by2, bw2 = 5.3, 0.35, 4.1
# Bottom mesh layer
mesh_r = plt.Rectangle((bx2, by2), bw2, 0.1, fc=MESH_COLOR, ec='#16a085',
                         lw=1.5, zorder=4, hatch='///', alpha=0.9)
ax3.add_patch(mesh_r)
ax3.text(bx2+bw2+0.1, by2+0.05, 'alumina mesh\n(500–1000 µm)', va='center',
         fontsize=7.5, color=MESH_COLOR)
# Powder on top of mesh
powder_m = plt.Rectangle((bx2, by2+0.1), bw2, 0.4, fc=POWDER_BROWN,
                           ec=BOAT_GRAY, lw=1.0, zorder=3, alpha=0.85)
ax3.add_patch(powder_m)
for xi in np.linspace(5.45, 9.25, 10):
    ax3.annotate('', xy=(xi, by2+0.08), xytext=(xi, by2-0.08),
                arrowprops=dict(arrowstyle='->', color=AR_BLUE, lw=0.9))
    ax3.annotate('', xy=(xi, by2+0.52), xytext=(xi, by2+0.12),
                arrowprops=dict(arrowstyle='->', color=AR_BLUE, lw=0.9,
                               alpha=0.5))
ax3.text(7.35, 0.62, 'vapor rises uniformly\nthrough entire bed', ha='center',
         va='center', fontsize=8, color='white', fontweight='bold')
arrow_ar(ax3, 4.6, 5.3, 0.75, label='Ar+Te')
for xi in np.linspace(5.45, 9.25, 6):
    ax3.annotate('', xy=(xi, 1.2), xytext=(xi, 0.97),
                arrowprops=dict(arrowstyle='->', color='#95a5a6', lw=0.8))


# ─────────────────────────────────────────────────
# PANEL 4 — Two-zone Te source control
# ─────────────────────────────────────────────────
ax4 = fig.add_axes([0.04, 0.12, 0.55, 0.175])
ax4.set_xlim(0, 8); ax4.set_ylim(0, 1.8); ax4.axis('off')

section_title(ax4, 0.05, 1.75, '4  Two-Zone Te Source Control',
              'Independent temperature zones set Te vapor pressure precisely')

# Zone 1 - Te source
z1 = FancyBboxPatch((0.3, 0.35), 2.0, 0.9, boxstyle="round,pad=0.05",
                     fc='#d2b4de', ec=TE_PURPLE, lw=2)
ax4.add_patch(z1)
ax4.text(1.3, 1.3, 'Zone 1: Te source', ha='center', va='bottom',
         fontsize=9, fontweight='bold', color=TE_PURPLE)
ax4.text(1.3, 1.15, 'T₁ = 300–450 °C', ha='center', va='bottom',
         fontsize=8, color=TE_PURPLE)
te_pellets = plt.Rectangle((0.55, 0.42), 0.55, 0.35, fc=TE_PURPLE,
                             ec='#6c3483', lw=1, zorder=3, alpha=0.8)
ax4.add_patch(te_pellets)
ax4.text(0.825, 0.60, 'Te', ha='center', va='center', fontsize=10,
         color='white', fontweight='bold')

# Zone 2 - carbon
z2 = FancyBboxPatch((3.0, 0.35), 3.5, 0.9, boxstyle="round,pad=0.05",
                     fc=FURNACE_LIGHT, ec=FURNACE_RED, lw=2)
ax4.add_patch(z2)
ax4.text(4.75, 1.3, 'Zone 2: Carbon bed', ha='center', va='bottom',
         fontsize=9, fontweight='bold', color=FURNACE_RED)
ax4.text(4.75, 1.15, 'T₂ = 700–900 °C', ha='center', va='bottom',
         fontsize=8, color=FURNACE_RED)
powder_boat(ax4, 3.4, 0.42, 2.5, 0.2, alpha=0.85)

# Arrow between zones
ax4.annotate('', xy=(3.0, 0.82), xytext=(2.3, 0.82),
             arrowprops=dict(arrowstyle='->', color=TE_PURPLE, lw=2.0))
ax4.text(2.65, 0.95, 'Te vapor', ha='center', va='bottom',
         fontsize=8, color=TE_PURPLE, fontweight='bold')

arrow_ar(ax4, 0.0, 0.3, 0.82, label='Ar')
ax4.annotate('', xy=(7.4, 0.82), xytext=(6.5, 0.82),
             arrowprops=dict(arrowstyle='->', color='#7f8c8d', lw=1.5))
ax4.text(7.5, 0.82, 'exhaust', va='center', fontsize=8, color='#7f8c8d')

# Inset: T1 vs vapor pressure
ax4i = fig.add_axes([0.37, 0.135, 0.18, 0.13])
T_vals = np.array([350, 380, 410, 440])
P_vals = np.array([0.15, 0.35, 0.65, 1.0])
ax4i.plot(T_vals, P_vals, 'o-', color=TE_PURPLE, lw=2, ms=5)
ax4i.set_xlabel('T₁ (°C)', fontsize=8)
ax4i.set_ylabel('Te vapor\npressure (a.u.)', fontsize=7)
ax4i.set_title('T₁ controls loading', fontsize=8, fontweight='bold')
ax4i.tick_params(labelsize=7)
ax4i.set_facecolor('#fdf9ff')
ax4i.spines[['top','right']].set_visible(False)


# ─────────────────────────────────────────────────
# PANEL 5 — Ar flow optimization
# ─────────────────────────────────────────────────
ax5 = fig.add_axes([0.62, 0.12, 0.36, 0.175])
ax5.set_xlim(0, 5); ax5.set_ylim(0, 1.8); ax5.axis('off')

section_title(ax5, 0.05, 1.75, '5  Ar Flow Rate Optimization',
              'Balance Te delivery vs. residence time')

flow_rates = [50, 100, 200]
outcomes = ['Depletes\nat inlet', 'Uniform\n✓', 'Blows through\nbefore deposit']
colors_f = ['#e74c3c', '#27ae60', '#e74c3c']
ys_f = [1.35, 0.85, 0.35]

for fr, out, col, yf in zip(flow_rates, outcomes, colors_f, ys_f):
    # mini tube
    tube_r = plt.Rectangle((0.5, yf-0.08), 3.5, 0.28,
                             fc=BG_WHITE, ec='#aaa', lw=1.2, zorder=2)
    ax5.add_patch(tube_r)
    # powder boat
    blen = 1.8
    powder_r = plt.Rectangle((1.5, yf-0.04), blen, 0.18,
                               fc=POWDER_BROWN, ec=BOAT_GRAY, lw=1.0,
                               zorder=3, alpha=0.8)
    ax5.add_patch(powder_r)
    # arrows (more arrows = more flow)
    n_arr = max(1, fr//60)
    for ai in range(n_arr):
        xstart = 0.6 + ai*0.18
        ax5.annotate('', xy=(xstart+0.25, yf+0.06),
                     xytext=(xstart, yf+0.06),
                     arrowprops=dict(arrowstyle='->', color=AR_BLUE,
                                    lw=1.5))
    ax5.text(0.35, yf+0.06, f'{fr}\nsccm', ha='center', va='center',
             fontsize=7.5, color=AR_BLUE, fontweight='bold')
    ax5.text(3.75, yf+0.06, out, ha='left', va='center',
             fontsize=8, color=col, fontweight='bold')

ax5.text(2.5, 0.03, '→  Run 3 flow-rate tests to find your optimum',
         ha='center', va='bottom', fontsize=8, color='#7f8c8d', style='italic')


# ─────────────────────────────────────────────────
# Overall title
# ─────────────────────────────────────────────────
fig.text(0.5, 0.995, 'Tube Furnace Scale-Up Strategies for Te-CVD on Carbon',
         ha='center', va='top', fontsize=16, fontweight='bold', color='#2c3e50')

plt.savefig('tube_furnace_upscale.png', dpi=150, bbox_inches='tight',
            facecolor=fig.get_facecolor())
print('Saved: tube_furnace_upscale.png')
