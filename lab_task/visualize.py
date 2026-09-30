"""
eCourts Uttarakhand - Data Visualization
=========================================
Generates 10 charts from ecourts_dataset.xlsx and saves them as PNG images.

Run:  python visualize.py
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
import numpy as np
import re
from pathlib import Path
from wordcloud import WordCloud
from collections import Counter

# ─────────────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────────────
print("Loading ecourts_dataset.xlsx ...")
df = pd.read_excel("ecourts_dataset.xlsx", sheet_name="Case Data")

# ── Parse key columns ─────────────────────────────────────────────────

# Extract Case Type, Case Number, Case Year from combined column
col = "Case Type/Case Number/Case Year"
df["Case Code"]   = df[col].str.extract(r"^([A-Z0-9]+)")                   # e.g. BA1
df["Case Year"]   = df[col].str.extract(r"/(\d{4})$").astype(float)        # e.g. 2026
df["Case Number"] = df[col].str.extract(r"/(\d+)/\d{4}$")                  # e.g. 1570

# Full case type label from _search column
df["Search Label"] = df["_search"].astype(str).str.strip("'\" ")

# Split petitioner vs respondent
party_col = "Petitioner Name Versus Respondent Name"
df["Petitioner"] = df[party_col].str.split(r"\\nVersus\\n|Versus").str[0].str.strip()
df["Respondent"] = df[party_col].str.split(r"\\nVersus\\n|Versus").str[-1].str.strip()

print(f"  {len(df)} rows loaded.")
print(f"  Search labels: {df['Search Label'].unique()}")
print(f"  Year range:    {df['Case Year'].min():.0f} - {df['Case Year'].max():.0f}")
print()

# ── Style ─────────────────────────────────────────────────────────────
plt.rcParams.update({
    "font.family":    "DejaVu Sans",
    "figure.dpi":     150,
    "axes.spines.top":    False,
    "axes.spines.right":  False,
    "axes.titleweight":   "bold",
    "axes.titlesize":     14,
    "axes.labelsize":     11,
})
PALETTE  = ["#1F3864","#2E75B6","#5BA3D9","#A8C8E8","#F4A261","#E76F51",
            "#2A9D8F","#E9C46A","#264653","#F4E1D2"]
OUT_DIR  = Path("charts")
OUT_DIR.mkdir(exist_ok=True)

saved = []

def save(name, fig):
    path = OUT_DIR / name
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    saved.append(str(path))
    print(f"  Saved: {path}")


# =============================================================================
# CHART 1 - HORIZONTAL BAR: Cases by Case Type (Search Label)
# =============================================================================
print("\n[1/10] Cases by Case Type (Bar Chart) ...")
counts = df["Search Label"].value_counts()
fig, ax = plt.subplots(figsize=(12, max(4, len(counts)*0.5)))
bars = ax.barh(counts.index, counts.values,
               color=PALETTE[:len(counts)], edgecolor="white", linewidth=0.5)
for bar, val in zip(bars, counts.values):
    ax.text(val + 5, bar.get_y() + bar.get_height()/2,
            str(val), va="center", fontsize=10)
ax.set_xlabel("Number of Cases")
ax.set_title("Number of Cases by Case Type")
ax.invert_yaxis()
ax.grid(axis="x", alpha=0.3)
fig.tight_layout()
save("01_cases_by_type_bar.png", fig)


# =============================================================================
# CHART 2 - PIE CHART: Distribution of Case Types
# =============================================================================
print("[2/10] Pie chart - Case Type distribution ...")
fig, ax = plt.subplots(figsize=(10, 7))
wedges, texts, autotexts = ax.pie(
    counts.values,
    labels=None,
    autopct="%1.1f%%",
    colors=PALETTE[:len(counts)],
    startangle=140,
    pctdistance=0.82,
    wedgeprops=dict(edgecolor="white", linewidth=1.5)
)
for at in autotexts:
    at.set_fontsize(9)
ax.legend(wedges, counts.index, loc="lower center",
          bbox_to_anchor=(0.5, -0.15), ncol=2, fontsize=9,
          framealpha=0.9)
ax.set_title("Distribution of Case Types", pad=20)
save("02_case_type_pie.png", fig)


# =============================================================================
# CHART 3 - BAR: Cases Filed Per Year
# =============================================================================
print("[3/10] Cases filed per year (Bar Chart) ...")
year_counts = df["Case Year"].dropna().astype(int).value_counts().sort_index()
fig, ax = plt.subplots(figsize=(12, 5))
bars = ax.bar(year_counts.index.astype(str), year_counts.values,
              color=PALETTE[1], edgecolor="white", width=0.65)
for bar, val in zip(bars, year_counts.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
            str(val), ha="center", fontsize=9)
ax.set_xlabel("Year")
ax.set_ylabel("Number of Cases")
ax.set_title("Cases Filed Per Year")
ax.grid(axis="y", alpha=0.3)
plt.xticks(rotation=45)
fig.tight_layout()
save("03_cases_per_year_bar.png", fig)


# =============================================================================
# CHART 4 - AREA CHART: Cumulative Cases Over Years
# =============================================================================
print("[4/10] Cumulative cases over years (Area Chart) ...")
cumulative = year_counts.cumsum()
fig, ax = plt.subplots(figsize=(12, 5))
ax.fill_between(cumulative.index.astype(str), cumulative.values,
                color=PALETTE[2], alpha=0.4, label="Cumulative")
ax.plot(cumulative.index.astype(str), cumulative.values,
        color=PALETTE[0], linewidth=2.5, marker="o", markersize=5)
ax.set_xlabel("Year")
ax.set_ylabel("Cumulative Case Count")
ax.set_title("Cumulative Cases Filed Over Years")
ax.grid(alpha=0.3)
plt.xticks(rotation=45)
fig.tight_layout()
save("04_cumulative_area.png", fig)


# =============================================================================
# CHART 5 - STACKED BAR: Case Type Breakdown per Year
# =============================================================================
print("[5/10] Case type breakdown per year (Stacked Bar) ...")
pivot = df.dropna(subset=["Case Year"])
pivot["Case Year"] = pivot["Case Year"].astype(int)
pivot_table = (pivot.groupby(["Case Year","Search Label"])
               .size().unstack(fill_value=0))
fig, ax = plt.subplots(figsize=(13, 6))
pivot_table.plot(kind="bar", stacked=True, ax=ax,
                 color=PALETTE[:len(pivot_table.columns)],
                 edgecolor="white", linewidth=0.5, width=0.75)
ax.set_xlabel("Year")
ax.set_ylabel("Number of Cases")
ax.set_title("Case Type Breakdown Per Year (Stacked Bar)")
ax.legend(loc="upper left", fontsize=8, bbox_to_anchor=(1,1))
plt.xticks(rotation=45)
ax.grid(axis="y", alpha=0.3)
fig.tight_layout()
save("05_stacked_bar_year_type.png", fig)


# =============================================================================
# CHART 6 - HEATMAP: Cases by Case Type x Year
# =============================================================================
print("[6/10] Heatmap - Case Type vs Year ...")
if len(pivot_table) > 1:
    fig, ax = plt.subplots(figsize=(max(8, len(pivot_table.columns)*2),
                                    max(5, len(pivot_table)*0.6)))
    sns.heatmap(pivot_table.T,
                annot=True, fmt="d", cmap="Blues",
                linewidths=0.5, linecolor="white",
                ax=ax, cbar_kws={"label": "Cases"})
    ax.set_title("Heatmap: Cases by Case Type and Year")
    ax.set_xlabel("Year")
    ax.set_ylabel("Case Type")
    fig.tight_layout()
    save("06_heatmap_type_year.png", fig)
else:
    print("  (skipped - not enough year variation)")


# =============================================================================
# CHART 7 - BOX PLOT: Case Number Distribution per Type
# =============================================================================
print("[7/10] Box plot - Case number distribution ...")
df["Case Number Num"] = pd.to_numeric(df["Case Number"], errors="coerce")
box_data = df.dropna(subset=["Case Number Num"])
if not box_data.empty:
    labels_order = (box_data.groupby("Search Label")["Case Number Num"]
                    .median().sort_values(ascending=False).index.tolist())
    fig, ax = plt.subplots(figsize=(13, 6))
    data_groups = [box_data[box_data["Search Label"]==l]["Case Number Num"].values
                   for l in labels_order]
    bp = ax.boxplot(data_groups, labels=labels_order, patch_artist=True,
                    medianprops=dict(color="white", linewidth=2),
                    boxprops=dict(linewidth=1.2),
                    whiskerprops=dict(linewidth=1.2),
                    capprops=dict(linewidth=1.2))
    for patch, color in zip(bp["boxes"], PALETTE[:len(labels_order)]):
        patch.set_facecolor(color)
    ax.set_ylabel("Case Number")
    ax.set_title("Box Plot: Case Number Range by Case Type")
    ax.grid(axis="y", alpha=0.3)
    plt.xticks(rotation=30, ha="right")
    fig.tight_layout()
    save("07_boxplot_case_numbers.png", fig)


# =============================================================================
# CHART 8 - SCATTER PLOT: Case Number vs Year (coloured by type)
# =============================================================================
print("[8/10] Scatter plot - Case Number vs Year ...")
sc_data = df.dropna(subset=["Case Number Num","Case Year"]).copy()
sc_data["Case Year"] = sc_data["Case Year"].astype(int)
if not sc_data.empty:
    labels_list = sc_data["Search Label"].unique()
    color_map   = {l: PALETTE[i % len(PALETTE)] for i,l in enumerate(labels_list)}
    fig, ax = plt.subplots(figsize=(13, 6))
    for label in labels_list:
        sub = sc_data[sc_data["Search Label"]==label]
        ax.scatter(sub["Case Year"], sub["Case Number Num"],
                   label=label, color=color_map[label],
                   alpha=0.5, s=20, edgecolors="none")
    ax.set_xlabel("Year")
    ax.set_ylabel("Case Number")
    ax.set_title("Scatter Plot: Case Numbers Over Years by Case Type")
    ax.legend(fontsize=8, bbox_to_anchor=(1,1), loc="upper left")
    ax.grid(alpha=0.2)
    plt.xticks(rotation=45)
    fig.tight_layout()
    save("08_scatter_year_casenumber.png", fig)


# =============================================================================
# CHART 9 - HORIZONTAL BAR: Top 20 Petitioners
# =============================================================================
print("[9/10] Top 20 petitioners (Bar Chart) ...")
pet = (df["Petitioner"]
       .dropna()
       .str.upper()
       .str.strip()
       .replace("", np.nan)
       .dropna()
       .value_counts()
       .head(20))
if not pet.empty:
    fig, ax = plt.subplots(figsize=(11, 7))
    bars = ax.barh(pet.index[::-1], pet.values[::-1],
                   color=PALETTE[0], edgecolor="white")
    for bar, val in zip(bars, pet.values[::-1]):
        ax.text(val + 0.2, bar.get_y() + bar.get_height()/2,
                str(val), va="center", fontsize=9)
    ax.set_xlabel("Number of Cases")
    ax.set_title("Top 20 Most Frequent Petitioners")
    ax.grid(axis="x", alpha=0.3)
    fig.tight_layout()
    save("09_top_petitioners_bar.png", fig)


# =============================================================================
# CHART 10 - WORD CLOUD: All Party Names
# =============================================================================
print("[10/10] Word Cloud - Party names ...")
all_names = " ".join(
    df["Petitioner"].dropna().astype(str).tolist() +
    df["Respondent"].dropna().astype(str).tolist()
)
# Remove filler words
stopwords = {"state","uttarakhand","versus","of","the","and","vs",
             "nan","none","viewdetails","for","case","number"}
wc = WordCloud(width=1400, height=700,
               background_color="white",
               colormap="Blues",
               stopwords=stopwords,
               max_words=150,
               prefer_horizontal=0.85)
wc.generate(all_names)
fig, ax = plt.subplots(figsize=(14, 7))
ax.imshow(wc, interpolation="bilinear")
ax.axis("off")
ax.set_title("Word Cloud: Most Common Party Names in Cases", fontsize=16, pad=15)
fig.tight_layout()
save("10_wordcloud_parties.png", fig)


# =============================================================================
# SUMMARY
# =============================================================================
print()
print("=" * 55)
print("  ALL CHARTS SAVED!")
print(f"  Folder: {OUT_DIR.resolve()}")
print("=" * 55)
print()
print("  Charts generated:")
for i, p in enumerate(saved, 1):
    print(f"  {i:2}. {Path(p).name}")
print()
print("  Open the 'charts' folder to see all PNG images.")
