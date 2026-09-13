"""
Generates docs/screenshots/dashboard-preview.png — a composed preview image of
the dashboard's hero, KPI band and two of its real chart panels, built from
the exact same data as the live dashboard. This is a rendered preview graphic
(not a raw browser screenshot) used for the README hero image.
"""
import json
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = "/usr/share/fonts/truetype/dejavu/"
def font(path, size):
    return ImageFont.truetype(FONT_DIR + path, size)

F_BRAND = font("DejaVuSans-Bold.ttf", 20)
F_NAV = font("DejaVuSans.ttf", 15)
F_EYEBROW = font("DejaVuSans-Bold.ttf", 15)
F_H1 = font("DejaVuSans-Bold.ttf", 40)
F_LEDE = font("DejaVuSans.ttf", 18)
F_KPI_LABEL = font("DejaVuSans.ttf", 14)
F_KPI_VALUE = font("DejaVuSans-Bold.ttf", 34)
F_KPI_SUB = font("DejaVuSans-Oblique.ttf", 13)
F_PANEL_TITLE = font("DejaVuSans-Bold.ttf", 17)
F_FOOT = font("DejaVuSans.ttf", 13)

NAVY = (14, 36, 56)
INK = (22, 33, 43)
PETROL = (11, 110, 127)
PETROL_LIGHT = (95, 169, 180)
CHARGE = (143, 209, 63)
AMBER = (226, 166, 61)
CORAL = (217, 99, 75)
GREY = (107, 116, 128)
LINE = (228, 233, 237)
PANEL = (246, 248, 249)
WHITE = (255, 255, 255)

W, H = 1600, 1980
img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)

def text_center_y(draw, xy, text, fnt, fill, anchor_left=True):
    draw.text(xy, text, font=fnt, fill=fill)

MARGIN = 64

# ---------------- Top bar ----------------
TOPBAR_H = 64
d.rectangle([0, 0, W, TOPBAR_H], fill=WHITE)
d.line([0, TOPBAR_H, W, TOPBAR_H], fill=LINE, width=2)
d.ellipse([MARGIN, 20, MARGIN + 24, 44], fill=CHARGE)
d.text((MARGIN + 34, 22), "Global EV Market Pulse", font=F_BRAND, fill=NAVY)
nav_items = ["World Map", "Countries", "Segmentation", "Trend", "Manufacturing", "Batteries", "2035 Outlook", "Sources"]
nx = W - MARGIN
for item in reversed(nav_items):
    tw = d.textlength(item, font=F_NAV)
    nx -= tw
    d.text((nx, 24), item, font=F_NAV, fill=GREY)
    nx -= 26

# ---------------- Hero ----------------
hy = TOPBAR_H + 40
d.text((MARGIN, hy), "WORLD ELECTRIC-VEHICLE INTELLIGENCE — 2026 EDITION", font=F_EYEBROW, fill=PETROL)
hy += 32
d.text((MARGIN, hy), "Which countries are actually winning the EV race —", font=F_H1, fill=NAVY)
hy += 50
d.text((MARGIN, hy), "and by how much.", font=F_H1, fill=NAVY)
hy += 56
lede = "A world-level view of EV production, sales, trade flows, battery supply chains and the"
lede2 = "road to 2035 — built entirely on the IEA's Global EV Outlook 2026."
d.text((MARGIN, hy), lede, font=F_LEDE, fill=GREY)
hy += 26
d.text((MARGIN, hy), lede2, font=F_LEDE, fill=GREY)
hy += 40

# ---------------- KPI band ----------------
kpis = [
    ("Global EV sales, 2025", "20.4M", "+20% year-on-year", False),
    ("Global sales share", "25%", "1 in 4 new cars sold worldwide", True),
    ("China's share of world sales", "\u224855%", "13M+ cars, world's largest market", False),
    ("EV fleet by 2035 (forecast)", "510M", "\u22486\u00d7 today, IEA no-new-policy case", True),
]
kpi_top = hy
kpi_h = 128
kpi_w = (W - 2 * MARGIN) / 4
d.rectangle([MARGIN, kpi_top, W - MARGIN, kpi_top + kpi_h], outline=LINE, width=2)
for i, (label, value, sub, accent) in enumerate(kpis):
    x0 = MARGIN + i * kpi_w
    if i > 0:
        d.line([x0, kpi_top, x0, kpi_top + kpi_h], fill=LINE, width=2)
    pad = 26
    d.text((x0 + pad, kpi_top + 20), label, font=F_KPI_LABEL, fill=GREY)
    d.text((x0 + pad, kpi_top + 44), value, font=F_KPI_VALUE, fill=(PETROL if accent else NAVY))
    d.text((x0 + pad, kpi_top + 92), sub, font=F_KPI_SUB, fill=GREY)
hy = kpi_top + kpi_h + 46

# ---------------- Charts (rendered with matplotlib, pasted in) ----------------
def render_leaderboard_chart():
    countries = ["China", "Germany", "United Kingdom", "T\u00fcrkiye", "United States",
                 "Korea", "Thailand", "Brazil", "Japan", "India", "Uruguay", "Egypt"]
    sales = [13_000_000, 850_000, 500_000, 240_000, 1_500_000, 200_000, 140_000,
             180_000, 100_000, 165_000, 13_500, 7_900]
    order = sorted(range(len(sales)), key=lambda i: sales[i])
    countries = [countries[i] for i in order]
    sales = [sales[i] for i in order]
    colors = [tuple(c/255 for c in CORAL) if c == "China" else tuple(c/255 for c in PETROL) for c in countries]
    fig, ax = plt.subplots(figsize=(8.6, 4.6), dpi=150)
    fig.patch.set_alpha(0)
    ax.set_facecolor((0, 0, 0, 0))
    bars = ax.barh(countries, sales, color=colors, height=0.62)
    ax.set_xscale("log")
    ax.set_xlim(1000, 3e7)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.spines["bottom"].set_color(tuple(c/255 for c in LINE))
    ax.tick_params(axis="y", labelsize=12, colors=tuple(c/255 for c in INK), length=0)
    ax.tick_params(axis="x", labelsize=0, length=0)
    ax.grid(axis="x", color=tuple(c/255 for c in LINE), linewidth=0.8)
    ax.set_axisbelow(True)
    for bar, val in zip(bars, sales):
        label = f"{val/1e6:.1f}M" if val >= 1e6 else f"{val/1e3:.0f}k"
        ax.text(bar.get_width()*1.15, bar.get_y()+bar.get_height()/2, label,
                va="center", fontsize=10.5, color=tuple(c/255 for c in INK))
    plt.tight_layout()
    fig.savefig("/tmp/chart_leaderboard.png", transparent=True)
    plt.close(fig)

def render_donut(labels, values, colors, path):
    fig, ax = plt.subplots(figsize=(3.4, 3.0), dpi=150)
    fig.patch.set_alpha(0)
    wedges, _ = ax.pie(values, colors=[tuple(c/255 for c in col) for col in colors],
                        startangle=90, counterclock=False,
                        wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
    for w, v in zip(wedges, values):
        ang = (w.theta2 + w.theta1) / 2
        import numpy as np
        x = 0.78 * np.cos(np.radians(ang))
        y = 0.78 * np.sin(np.radians(ang))
        ax.text(x, y, f"{v}%", ha="center", va="center", fontsize=12, fontweight="bold",
                color="white" if v > 15 else tuple(c/255 for c in INK))
    plt.tight_layout(pad=0.2)
    fig.savefig(path, transparent=True)
    plt.close(fig)

def render_trend_chart():
    years = [2020, 2021, 2022, 2023, 2024, 2025, 2026]
    sales = [3.1, 6.6, 10.5, 14.2, 17.1, 20.4, 23.0]
    is_forecast = [False]*6 + [True]
    fig, ax = plt.subplots(figsize=(4.2, 2.6), dpi=150)
    fig.patch.set_alpha(0)
    ax.set_facecolor((0, 0, 0, 0))
    colors = [tuple(c/255 for c in (AMBER if f else PETROL)) for f in is_forecast]
    ax.plot(years, sales, color=tuple(c/255 for c in PETROL), linewidth=2.2, zorder=1)
    ax.scatter(years, sales, color=colors, s=36, zorder=2)
    ax.fill_between(years, sales, color=tuple(c/255 for c in PETROL), alpha=0.08)
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color(tuple(c/255 for c in LINE))
    ax.tick_params(labelsize=9, colors=tuple(c/255 for c in INK), length=0)
    ax.grid(axis="y", color=tuple(c/255 for c in LINE), linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_ylabel("Million units", fontsize=9, color=tuple(c/255 for c in GREY))
    plt.tight_layout()
    fig.savefig("/tmp/chart_trend.png", transparent=True)
    plt.close(fig)

def render_recycling_chart():
    years = [2025, 2030, 2037]
    gwh = [45.5, 330, 1430]
    fig, ax = plt.subplots(figsize=(4.2, 2.6), dpi=150)
    fig.patch.set_alpha(0)
    ax.set_facecolor((0, 0, 0, 0))
    ax.plot(years, gwh, color=tuple(c/255 for c in CORAL), linewidth=2.2, marker="o", markersize=6)
    ax.fill_between(years, gwh, color=tuple(c/255 for c in CORAL), alpha=0.08)
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color(tuple(c/255 for c in LINE))
    ax.tick_params(labelsize=9, colors=tuple(c/255 for c in INK), length=0)
    ax.grid(axis="y", color=tuple(c/255 for c in LINE), linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_ylabel("GWh (EOL batteries)", fontsize=9, color=tuple(c/255 for c in GREY))
    for x, y in zip(years, gwh):
        ax.annotate(f"{y:,.0f}", (x, y), textcoords="offset points", xytext=(0, 8),
                    ha="center", fontsize=8.5, color=tuple(c/255 for c in INK))
    plt.tight_layout()
    fig.savefig("/tmp/chart_recycling.png", transparent=True)
    plt.close(fig)

render_leaderboard_chart()
render_donut(["BEV", "PHEV/EREV"], [65, 35], [CHARGE, AMBER], "/tmp/chart_powertrain.png")
render_donut(["China", "Europe", "N. America", "Other"], [60, 15, 15, 10],
             [CORAL, PETROL, NAVY, GREY], "/tmp/chart_origin.png")
render_trend_chart()
render_recycling_chart()

# ---------------- Panels ----------------
panel_top = hy
panel_h = 560
left_w = int((W - 2*MARGIN) * 0.6) - 14
right_w = (W - 2*MARGIN) - left_w - 28

def panel(x0, y0, w, h, title):
    d.rectangle([x0, y0, x0+w, y0+h], fill=PANEL, outline=LINE, width=2)
    d.text((x0+22, y0+18), title, font=F_PANEL_TITLE, fill=NAVY)

panel(MARGIN, panel_top, left_w, panel_h, "Top Electric-Car Markets by 2025 Sales")
chart_img = Image.open("/tmp/chart_leaderboard.png").convert("RGBA")
scale = (left_w - 60) / chart_img.width
chart_img = chart_img.resize((int(chart_img.width*scale), int(chart_img.height*scale)))
img.paste(chart_img, (MARGIN + 30, panel_top + 70), chart_img)

rx = MARGIN + left_w + 28
panel(rx, panel_top, right_w, panel_h, "Market Segmentation, 2025")

col_w = (right_w - 60) // 2
p1 = Image.open("/tmp/chart_powertrain.png").convert("RGBA")
p2 = Image.open("/tmp/chart_origin.png").convert("RGBA")
pscale = col_w / p1.width
p1r = p1.resize((int(p1.width*pscale), int(p1.height*pscale)))
p2r = p2.resize((int(p2.width*pscale), int(p2.height*pscale)))

cap_y = panel_top + 74
img.paste(p1r, (rx + 20, cap_y + 26), p1r)
img.paste(p2r, (rx + 30 + col_w, cap_y + 26), p2r)
d.text((rx + 20, cap_y), "Powertrain mix", font=F_KPI_LABEL, fill=GREY)
d.text((rx + 30 + col_w, cap_y), "Manufacturer origin", font=F_KPI_LABEL, fill=GREY)

legend_y = cap_y + 26 + p1r.height + 20
legend_items_1 = [("BEV — 65%", CHARGE), ("PHEV/EREV — 35%", AMBER)]
legend_items_2 = [("China — 60%", CORAL), ("Europe — 15%", PETROL), ("N. America — 15%", NAVY), ("Other — 10%", GREY)]
ly = legend_y
for label, col in legend_items_1:
    d.ellipse([rx+20, ly, rx+32, ly+12], fill=col)
    d.text((rx+40, ly-2), label, font=F_KPI_SUB, fill=INK)
    ly += 20
ly2 = legend_y
for label, col in legend_items_2:
    lx = rx + 30 + col_w
    d.ellipse([lx, ly2, lx+12, ly2+12], fill=col)
    d.text((lx+20, ly2-2), label, font=F_KPI_SUB, fill=INK)
    ly2 += 20

hy = panel_top + panel_h + 40

# ---------------- Row 2: Trend + Recycling outlook ----------------
panel2_top = hy
panel2_h = 500
half_w = (W - 2*MARGIN - 28) // 2

panel(MARGIN, panel2_top, half_w, panel2_h, "Global EV Sales, 2020\u20132026 (million)")
tchart = Image.open("/tmp/chart_trend.png").convert("RGBA")
tscale = (half_w - 60) / tchart.width
tchart_r = tchart.resize((int(tchart.width*tscale), int(tchart.height*tscale)))
img.paste(tchart_r, (MARGIN + 30, panel2_top + 60), tchart_r)

rx2 = MARGIN + half_w + 28
panel(rx2, panel2_top, half_w, panel2_h, "Global End-of-Life Battery Volume (GWh)")
rchart = Image.open("/tmp/chart_recycling.png").convert("RGBA")
rscale = (half_w - 60) / rchart.width
rchart_r = rchart.resize((int(rchart.width*rscale), int(rchart.height*rscale)))
img.paste(rchart_r, (rx2 + 30, panel2_top + 60), rchart_r)

hy = panel2_top + panel2_h + 40

# ---------------- Footer ----------------
d.line([MARGIN, hy, W-MARGIN, hy], fill=LINE, width=2)
hy += 22
d.text((MARGIN, hy), "Global EV Market Pulse — a data portfolio project by Milad Shabani.", font=F_FOOT, fill=GREY)
right_text = "Data: IEA, Global EV Outlook 2026 (CC BY 4.0)"
tw = d.textlength(right_text, font=F_FOOT)
d.text((W - MARGIN - tw, hy), right_text, font=F_FOOT, fill=PETROL)

crop_bottom = min(H, hy + 46)
img = img.crop((0, 0, W, crop_bottom))
img.save("dashboard-preview.png")
print("Saved dashboard-preview.png", img.size)
