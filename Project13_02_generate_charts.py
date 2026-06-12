"""
Energy Transition Charts — 6 Publication-Quality Charts
Using REAL data from sustainability reports
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import os

OUT = "/home/claude/energy_transition/charts/"
os.makedirs(OUT, exist_ok=True)

RELIANCE  = "#E65100"
TATA      = "#1565C0"
ITC       = "#2E7D32"
NESTLE    = "#6A1B9A"
BG        = "#FAFAFA"
GRID      = "#E0E0E0"
TEXT      = "#212121"
COMPANIES = ["Reliance\nIndustries", "Tata\nMotors", "ITC\nLimited", "Nestlé\nGlobal"]
COLORS    = [RELIANCE, TATA, ITC, NESTLE]

plt.rcParams.update({
    "font.family": "DejaVu Sans","font.size": 10,"axes.titlesize": 13,
    "axes.titleweight": "bold","axes.labelsize": 10,
    "axes.spines.top": False,"axes.spines.right": False,
    "axes.grid": True,"grid.color": GRID,"grid.linewidth": 0.7,
    "figure.facecolor": BG,"axes.facecolor": BG,
    "text.color": TEXT,"axes.labelcolor": TEXT,"xtick.color": TEXT,"ytick.color": TEXT,
})

def add_source(ax, text):
    ax.text(0, -0.13, text, transform=ax.transAxes, fontsize=7, color="#757575", style="italic", ha="left")

def add_watermark(fig):
    fig.text(0.98, 0.01, "Sanjana Sabat | ESG & Sustainability Analyst | Real Data from Official SRs",
             ha="right", va="bottom", fontsize=7, color="#BDBDBD")

# ── CHART 1: Scope 1+2 Comparison ──────────────────────────────────────────
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6), facecolor=BG)
fig.suptitle("Scope 1 & 2 GHG Emissions — Cross-Sector Comparison (FY2024-25)",
             fontsize=14, fontweight="bold", y=1.01)
scope1 = [36.46, 0.071, 1.105, 2.82]
scope2 = [1.47,  0.214, 0.159, 0.22]
x = np.arange(4); w = 0.35
ax1.bar(x-w/2, scope1, w, label="Scope 1 (Direct)", color=COLORS, alpha=0.88, edgecolor="white")
ax1.bar(x+w/2, scope2, w, label="Scope 2 MB", color=COLORS, alpha=0.50, edgecolor="white", hatch="///")
ax1.set_yscale("log"); ax1.set_xticks(x); ax1.set_xticklabels(COMPANIES, fontsize=9)
ax1.set_ylabel("MtCO₂e (log scale)"); ax1.set_title("S1 & S2 Absolute (Log Scale)")
ax1.legend(fontsize=8)
totals = [37.93, 0.285, 1.264, 3.04]
bars_t = ax2.barh(COMPANIES[::-1], totals[::-1], color=COLORS[::-1], alpha=0.85, edgecolor="white", height=0.55)
ax2.set_xscale("log"); ax2.set_xlabel("MtCO₂e (log scale)"); ax2.set_title("Total S1+S2 (Log Scale)")
for bar, val in zip(bars_t, totals[::-1]):
    ax2.text(val*1.3, bar.get_y()+bar.get_height()/2, f"{val:.3f} Mt", va="center", fontsize=8.5, fontweight="bold")
add_source(ax2, "Sources: Reliance BRSR 2024-25 | TM BRSR p.71 | ITC SR 2025 | Nestlé SR Appendix 4")
plt.tight_layout(pad=1.5); add_watermark(fig)
plt.savefig(f"{OUT}chart1_scope1_2_comparison.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(); print("✅ Chart 1")

# ── CHART 2: Renewable Energy Progress ──────────────────────────────────────
fig, ax = plt.subplots(figsize=(13, 7), facecolor=BG)
years_labels = ["FY22/23\n(Approx)", "FY23/24\n(Actual)", "FY24/25\n(Actual)", "Target\n(~2030)"]
re_data = {
    "Reliance Industries": [0.80, 0.95, 1.11, 100.0],
    "Tata Motors":         [18.0, 30.0, 45.0, 100.0],
    "ITC Limited":         [44.0, 48.0, 52.0, 100.0],
    "Nestlé Global":       [36.0, 38.5, 40.6, 100.0],
}
x = np.arange(4)
markers = ["o", "s", "D", "^"]
for (co, vals), color, mk in zip(re_data.items(), COLORS, markers):
    ax.plot(x[:3], vals[:3], color=color, lw=2.5, marker=mk, markersize=9, label=co, zorder=3)
    ax.plot(x[2:], vals[2:], color=color, lw=1.5, linestyle="--", marker=mk, markersize=7, alpha=0.6, zorder=2)
    ax.annotate(f"  {vals[2]:.1f}%", xy=(x[2], vals[2]), fontsize=9, color=color, fontweight="bold", va="center")
ax.axvline(2.5, color="#B0BEC5", linestyle=":", lw=1.5)
ax.axhline(100, color="#4CAF50", lw=1.2, linestyle="--", alpha=0.5)
ax.fill_betweenx([0,105], 2.5, 3.5, alpha=0.06, color="#9E9E9E")
ax.set_xticks(x); ax.set_xticklabels(years_labels, fontsize=10)
ax.set_ylabel("Renewable Energy Share (%)"); ax.set_ylim(0, 112); ax.set_yticks(range(0,105,10))
ax.yaxis.set_major_formatter(mticker.PercentFormatter())
ax.set_title("Renewable Energy Progress: FY2023 → FY2025 → 2030 Target")
ax.legend(loc="upper left", fontsize=9, framealpha=0.9, ncol=2)
ax.text(0.62, 0.15,
    "Key Insights:\n• Tata Motors: fastest growth (+15pp in FY25)\n"
    "• ITC: highest absolute level (52%)\n"
    "• Reliance: only 1.1% — justifies $10B+ RE investment\n"
    "• Nestlé mfg electricity: 98.6% RE (total energy basis shown here)",
    transform=ax.transAxes, fontsize=8, va="bottom",
    bbox=dict(boxstyle="round,pad=0.5", facecolor="#E8F5E9", edgecolor="#4CAF50", alpha=0.9))
add_source(ax, "Sources: Reliance BRSR 2024-25 | TM BRSR 2024-25 | ITC SR 2024-25 | Nestlé SR 2024")
plt.tight_layout(pad=1.5); add_watermark(fig)
plt.savefig(f"{OUT}chart2_renewable_energy_progress.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(); print("✅ Chart 2")

# ── CHART 3: Net Zero Timeline (Gantt) ──────────────────────────────────────
fig, ax = plt.subplots(figsize=(15, 7), facecolor=BG)
milestones = {
    "Reliance\nIndustries": {
        "color": RELIANCE, "y": 4, "nz": 2035,
        "events": [(2024,"1.11% RE | $10B+ New Energy"),(2027,"Green H2 plant (Jamnagar)"),(2030,"100 GW RE target"),(2035,"★ NET ZERO")]
    },
    "Tata Motors": {
        "color": TATA, "y": 3, "nz": 2045,
        "events": [(2024,"45% RE | ENCON 18,423 tCO₂ saved"),(2028,"RE100 target"),(2030,"Water Positive + ZWtL"),(2040,"★ Net Zero PV"),(2045,"★ Net Zero CV")]
    },
    "ITC Limited": {
        "color": ITC, "y": 2, "nz": 2050,
        "events": [(2024,"52% RE | Carbon Positive | 6.46 Mt seq."),(2028,"−50% GHG intensity vs FY19"),(2050,"★ NET ZERO")]
    },
    "Nestlé Global": {
        "color": NESTLE, "y": 1, "nz": 2050,
        "events": [(2024,"40.6% RE | Net: 65.74 Mt | SBTi ✓"),(2025,"FLAG −25% vs 2018"),(2030,"−50% GHG vs 2018"),(2050,"★ NET ZERO")]
    },
}
ax.set_xlim(2022, 2058); ax.set_ylim(0.3, 5.2)
ax.set_yticks([1,2,3,4]); ax.set_yticklabels(list(milestones.keys()), fontsize=11, fontweight="bold")
ax.set_xticks(range(2024, 2056, 2)); ax.set_xlabel("Year", fontsize=10)
ax.set_title("Net Zero Journey: Key Milestones & Commitments", fontsize=13, fontweight="bold")
for co, data in milestones.items():
    y, color, nz = data["y"], data["color"], data["nz"]
    ax.barh(y, nz-2024, left=2024, height=0.22, color=color, alpha=0.15, zorder=1)
    ax.axhline(y, color=color, alpha=0.25, lw=0.8)
    for i, (yr, lbl) in enumerate(data["events"]):
        is_nz = "NET ZERO" in lbl or "★" in lbl
        ax.scatter(yr, y, s=200 if is_nz else 70, color=color, zorder=5,
                   marker="*" if is_nz else "o", edgecolors="white", linewidth=1.2)
        off = 0.22 if i % 2 == 0 else -0.24
        ax.annotate(f"{yr}: {lbl}", xy=(yr, y), xytext=(yr, y+off),
                    fontsize=6.5, color=color, ha="center",
                    va="bottom" if off>0 else "top",
                    arrowprops=dict(arrowstyle="-", color=color, lw=0.7, alpha=0.5))
for yr, lbl in [(2030,"2030\nClimate Targets"),(2050,"2050\nNet Zero")]:
    ax.axvline(yr, color="#78909C", linestyle="--", lw=1.5, alpha=0.6)
    ax.text(yr+0.3, 0.42, lbl, fontsize=7.5, color="#546E7A", va="bottom")
add_source(ax, "Sources: All official sustainability reports / BRSR filings (FY2024-25)")
plt.tight_layout(pad=1.5); add_watermark(fig)
plt.savefig(f"{OUT}chart3_net_zero_timeline.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(); print("✅ Chart 3")

# ── CHART 4: Scope 3 Analysis ────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(14, 6), facecolor=BG)
# Tata Motors Scope 3 pie
tm_cats = ["Use of Sold\nProducts (152.6 Mt)","Purchased Goods\n& Services (9.4 Mt)",
           "Franchises\n(0.19 Mt)","Other Cat.\n(0.1 Mt)"]
tm_vals = [152.60, 9.407, 0.191, 0.102]
cat_colors = ["#1565C0","#1976D2","#42A5F5","#90CAF9"]
wedges, texts, autotexts = axes[0].pie(tm_vals, labels=None, colors=cat_colors,
    autopct=lambda p: f"{p:.1f}%" if p>1 else "", startangle=90, pctdistance=0.75,
    wedgeprops={"edgecolor":"white","linewidth":1.5})
for at in autotexts: at.set_fontsize(8)
axes[0].legend(wedges, tm_cats, loc="lower left", fontsize=7.5, bbox_to_anchor=(-0.1, -0.22), ncol=2)
axes[0].set_title(f"Tata Motors: Scope 3 by Category\n(Total: 162.3 MtCO₂e)", fontsize=11, fontweight="bold")
axes[0].text(0, 0, "162.3\nMtCO₂e", ha="center", va="center", fontsize=10, fontweight="bold", color=TATA)
# S1+2 vs S3 all companies
s12 = [37.93, 0.285, 1.264, 3.04]
s3  = [200.0, 162.30, 1.062, 66.01]
x = np.arange(4); w = 0.38
b1 = axes[1].bar(x-w/2, s12, w, label="Scope 1+2", color=COLORS, alpha=0.9, edgecolor="white")
b2 = axes[1].bar(x+w/2, s3,  w, label="Scope 3",   color=COLORS, alpha=0.45, edgecolor="white", hatch="///")
axes[1].set_yscale("log"); axes[1].set_xticks(x); axes[1].set_xticklabels(COMPANIES, fontsize=9)
axes[1].set_ylabel("MtCO₂e (log scale)"); axes[1].set_title("S1+2 vs Scope 3 (Log Scale)")
axes[1].legend(fontsize=9)
for b,v in zip(b1,s12): axes[1].text(b.get_x()+b.get_width()/2, v*1.4, f"{v:.2f}", ha="center", fontsize=7.5, fontweight="bold")
for b,v in zip(b2,s3):  axes[1].text(b.get_x()+b.get_width()/2, v*1.4, f"{v:.0f}", ha="center", fontsize=7.5, fontweight="bold")
axes[1].text(0.04, 0.04, "Reliance S3: estimated.\nAll others from official SRs.", transform=axes[1].transAxes, fontsize=7, color="#757575", style="italic")
fig.suptitle("Scope 3 Value Chain Emissions Analysis", fontsize=13, fontweight="bold", y=1.01)
add_source(axes[1], "TM BRSR p.79 | ITC SR 2024-25 | Nestlé SR App.4 | Reliance BRSR 2024-25")
plt.tight_layout(pad=2.0); add_watermark(fig)
plt.savefig(f"{OUT}chart4_scope3_analysis.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(); print("✅ Chart 4")

# ── CHART 5: Radar Scorecard ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 10), subplot_kw=dict(polar=True), facecolor=BG)
categories = ["RE Progress\n(vs target)","Net Zero\nAmbition","Scope 3\nStrategy",
              "Disclosure\nQuality","Carbon\nSequestration","Water\nStewardship",
              "Circular\nEconomy","SBTi\nAlignment"]
N = len(categories)
angles = np.linspace(0, 2*np.pi, N, endpoint=False).tolist() + [0]
scores = {
    "Reliance Industries": [1, 8, 4, 7, 2, 4, 3, 5],
    "Tata Motors":         [7, 8, 6, 9, 4, 8, 8, 7],
    "ITC Limited":         [8, 7, 7, 8, 10, 7, 7, 7],
    "Nestlé Global":       [7, 7, 8, 9, 6, 7, 7, 10],
}
for co, sc, color in zip(scores.keys(), scores.values(), COLORS):
    vals = sc + sc[:1]
    ax.plot(angles, vals, color=color, linewidth=2.2, label=co, zorder=3)
    ax.fill(angles, vals, color=color, alpha=0.10, zorder=2)
ax.set_xticks(angles[:-1]); ax.set_xticklabels(categories, fontsize=9, fontweight="bold")
ax.set_ylim(0,10); ax.set_yticks([2,4,6,8,10]); ax.set_yticklabels(["2","4","6","8","10"], fontsize=7)
ax.set_title("Decarbonisation Performance Scorecard\n(Analyst Assessment — Scale 0–10)", fontsize=13, fontweight="bold", pad=25)
ax.legend(loc="lower left", bbox_to_anchor=(-0.18,-0.14), fontsize=10, framealpha=0.9, ncol=2)
ax.text(-0.12,-0.14,"Scores: analyst assessment from SR content. RE Progress = current % vs 100% target.\nDisclosure Quality based on GHG Protocol alignment, assurance coverage, BRSR/GRI adherence.",
        transform=ax.transAxes, fontsize=7.5, color="#757575", style="italic")
add_watermark(fig)
plt.savefig(f"{OUT}chart5_scorecard_radar.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(); print("✅ Chart 5")

# ── CHART 6: GHG Intensity & Investment ─────────────────────────────────────
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6), facecolor=BG)
# Intensity index
intensity = [100.0, 0.75, 3.33, 8.01]
trends    = ["Stable","↓3.8%/vehicle","↓46% vs FY19","↓24.5% vs 2018"]
bars = ax1.bar(COMPANIES, intensity, color=COLORS, alpha=0.85, edgecolor="white", width=0.55)
ax1.set_ylabel("Relative GHG Intensity Index\n(Reliance = 100)", fontsize=9)
ax1.set_title("GHG Emission Intensity Index\n(Relative Scale — Lower is Better)", fontsize=11, fontweight="bold")
for b, v, tr in zip(bars, intensity, trends):
    ax1.text(b.get_x()+b.get_width()/2, b.get_height()+1.5, f"{v:.1f}\n{tr}", ha="center", fontsize=8.5, fontweight="bold", color=b.get_facecolor())
ax1.set_ylim(0, 115)
ax1.text(0.04, 0.96, "Lower = Better Performance", transform=ax1.transAxes, fontsize=8, color="#4CAF50",
         fontweight="bold", va="top", bbox=dict(boxstyle="round", facecolor="#E8F5E9", edgecolor="#4CAF50"))
add_source(ax1, "Relative index: Reliance=100 base. Others scaled by S1+S2 / revenue ratio.")
# Investment scale
inv_labels = ["Reliance\n$10B+\nNew Energy","Tata Motors\nRE100 + EV\nExpansion","ITC\nForestry +\nRE","Nestlé\nRegen. Agri +\nRE Transition"]
inv_scale  = [10.0, 3.5, 1.5, 2.8]
inv_cats   = ["Green H₂ + Solar\n+ Battery + Wind","76.5 MWp Solar\n+ PPA + iRECs","Agroforestry\n+ Solar Expansion","Regenerative Farming\n+ RE Procurement"]
bars2 = ax2.bar(inv_labels, inv_scale, color=COLORS, alpha=0.85, edgecolor="white", width=0.55)
ax2.set_ylabel("Approximate Investment Scale\n(USD Billion equivalent)", fontsize=9)
ax2.set_title("Energy Transition Investment Scale\n(Indicative from SR Disclosures)", fontsize=11, fontweight="bold")
for b, cat in zip(bars2, inv_cats):
    ax2.text(b.get_x()+b.get_width()/2, b.get_height()+0.1, cat, ha="center", fontsize=7, color=b.get_facecolor(), fontweight="bold")
ax2.set_ylim(0, 13)
ax2.text(0.04, 0.96, "Reliance $10B+ explicitly stated.\nOthers are indicative estimates.", transform=ax2.transAxes, fontsize=7.5, color="#757575", style="italic", va="top")
fig.suptitle("GHG Intensity & Energy Transition Investment Comparison", fontsize=13, fontweight="bold", y=1.01)
plt.tight_layout(pad=2.0); add_watermark(fig)
plt.savefig(f"{OUT}chart6_intensity_investment.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(); print("✅ Chart 6")

print("\n✅ ALL 6 CHARTS DONE")
