#!/usr/bin/env python3
"""Regenerate manuscript Figures 2–6 and Supplementary Figure S1.

The script has no dependency on the copyrighted source PDFs or on intermediate
CSV files. It reads only the reviewer-revised supplementary evidence workbook
and writes publication-ready PNG/PDF figures plus auditable derived tables.
Figure 1 (PRISMA flow diagram) is intentionally excluded because it is prepared
manually.

Example
-------
python generate_figures.py \
  --workbook ../Supplementary_Data_REV2.xlsx \
  --output-dir generated_figures
"""

from __future__ import annotations

import argparse
import ast
from collections import Counter
from pathlib import Path
import textwrap

import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.lines import Line2D
import numpy as np
import pandas as pd


NAVY = "#183B5B"
BLUE = "#3277A8"
TEAL = "#2A9D8F"
ORANGE = "#E88C4A"
GOLD = "#D9B44A"
RED = "#C9553D"
INK = "#263238"
MID = "#77838F"
LIGHT = "#E7ECF0"

plt.rcParams.update(
    {
        "font.family": "DejaVu Sans",
        "font.size": 9.2,
        "axes.titlesize": 12.2,
        "axes.labelsize": 9.5,
        "xtick.labelsize": 8.2,
        "ytick.labelsize": 8.2,
        "axes.titleweight": "bold",
        "axes.spines.top": False,
        "axes.spines.right": False,
        "figure.facecolor": "white",
        "axes.facecolor": "white",
        "savefig.facecolor": "white",
    }
)


REQUIRED_SHEETS = {
    "Article Evidence",
    "Contribution Analysis",
    "Outcome Evidence",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate manuscript Figures 2–6 and Supplementary Figure S1 from the evidence workbook."
    )
    parser.add_argument("--workbook", required=True, type=Path, help="Path to the evidence XLSX workbook.")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("generated_figures"),
        help="Output directory (default: generated_figures).",
    )
    parser.add_argument("--dpi", type=int, default=420, help="PNG resolution (default: 420 dpi).")
    return parser.parse_args()


def parse_list(value: object) -> list[str]:
    """Parse Excel cells containing Python-style lists or semicolon-separated values."""
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return []
    text = str(value).strip()
    if not text:
        return []
    if text.startswith("[") and text.endswith("]"):
        try:
            parsed = ast.literal_eval(text)
            if isinstance(parsed, (list, tuple, set)):
                return [str(x).strip() for x in parsed if str(x).strip()]
        except (SyntaxError, ValueError):
            pass
    return [part.strip() for part in text.split(";") if part.strip()]


def yes(value: object) -> bool:
    return str(value).strip().lower() == "yes"


def require_columns(frame: pd.DataFrame, sheet: str, columns: list[str]) -> None:
    missing = [column for column in columns if column not in frame.columns]
    if missing:
        raise ValueError(f"Sheet '{sheet}' is missing required column(s): {', '.join(missing)}")


def load_workbook(path: Path) -> dict[str, pd.DataFrame]:
    if not path.exists():
        raise FileNotFoundError(path)
    excel = pd.ExcelFile(path)
    missing_sheets = sorted(REQUIRED_SHEETS - set(excel.sheet_names))
    if missing_sheets:
        raise ValueError(f"Workbook is missing required sheet(s): {', '.join(missing_sheets)}")
    data = {
        sheet: pd.read_excel(path, sheet_name=sheet).fillna("")
        for sheet in REQUIRED_SHEETS
    }
    article = data["Article Evidence"]
    require_columns(
        article,
        "Article Evidence",
        [
            "Record_ID",
            "Year",
            "System_boundary_class",
            "Database_group",
            "Method_groups",
            "Technology_groups",
            "Contaminant_groups",
            "Setting_group",
            "Scope_group",
            "Normalization_type",
            "Weighting_type",
            "Contribution_analysis",
            "Sensitivity_analysis",
            "Uncertainty_analysis",
            "Circularity_waste_prevention",
            "Circularity_sediment_recycling_recovery",
            "Circularity_byproduct_or_secondary_input",
            "Circularity_avoided_burden_modelled",
            "Score_FU_Boundaries",
            "Score_Foreground_Data",
            "Score_Background_DB",
            "Score_LCIA_Reporting",
            "Score_Allocation_Cutoff",
            "Score_Robustness",
            "Descriptive_Profile_Sum_0_12",
        ],
    )
    require_columns(
        data["Contribution Analysis"],
        "Contribution Analysis",
        ["Record_ID", "Broad_process_groups"],
    )
    require_columns(
        data["Outcome Evidence"],
        "Outcome Evidence",
        ["Record_ID", "Impact_category", "Outcome_signal"],
    )
    for column in ["Method_groups", "Technology_groups", "Contaminant_groups"]:
        article[column] = article[column].map(parse_list)
    article["Year"] = pd.to_numeric(article["Year"], errors="coerce")
    return data


def panel_label(ax: plt.Axes, label: str, fontsize: float = 17) -> None:
    ax.text(
        -0.055,
        1.035,
        f"({label})",
        transform=ax.transAxes,
        fontsize=fontsize,
        fontweight="bold",
        color=NAVY,
        va="bottom",
        ha="left",
        clip_on=False,
    )


def clean_axis(ax: plt.Axes) -> None:
    ax.grid(False)
    ax.tick_params(colors=INK)
    ax.xaxis.label.set_color(INK)
    ax.yaxis.label.set_color(INK)
    ax.title.set_color(INK)
    ax.spines["left"].set_color("#9AA4AD")
    ax.spines["bottom"].set_color("#9AA4AD")


def save_figure(fig: plt.Figure, output_dir: Path, stem: str, dpi: int) -> tuple[Path, Path]:
    png = output_dir / f"{stem}.png"
    pdf = output_dir / f"{stem}.pdf"
    fig.savefig(png, dpi=dpi, bbox_inches="tight")
    fig.savefig(pdf, bbox_inches="tight")
    plt.close(fig)
    return png, pdf


def count_expanded(frame: pd.DataFrame, column: str) -> pd.DataFrame:
    counter: Counter[str] = Counter()
    for values in frame[column]:
        counter.update(set(values))
    return pd.DataFrame(counter.most_common(), columns=["Category", "Studies"])


def lollipop(ax: plt.Axes, frame: pd.DataFrame, total: int, color: str) -> None:
    data = frame.sort_values("Studies", ascending=True).reset_index(drop=True)
    y = np.arange(len(data))
    ax.hlines(y, 0, data["Studies"], color=LIGHT, lw=3, zorder=1)
    ax.scatter(data["Studies"], y, s=125, color=color, edgecolor="white", lw=1.0, zorder=2)
    for yy, value in zip(y, data["Studies"]):
        ax.text(value + 0.45, yy, f"{int(value)}  ({value / total:.0%})", va="center", color=INK, fontsize=11)
    ax.set_yticks(y, [textwrap.fill(str(x), 24) for x in data["Category"]])
    ax.set_xlim(0, max(data["Studies"]) * 1.34)
    ax.set_xlabel("Studies")
    ax.tick_params(axis="both", labelsize=12)
    ax.xaxis.label.set_size(12)
    clean_axis(ax)


def dot_strip(
    ax: plt.Axes,
    labels: list[str],
    counts: list[int],
    colors: list[str],
    total: int,
) -> None:
    rows = len(labels)
    xs = np.arange(total)
    for i, (label, count, color) in enumerate(zip(labels, counts, colors)):
        y = rows - 1 - i
        ax.scatter(xs, np.full(total, y), s=32, color=LIGHT, edgecolor="none")
        ax.scatter(xs[:count], np.full(count, y), s=40, color=color, edgecolor="white", lw=0.3)
        ax.text(-2.2, y, label, ha="right", va="center", fontsize=11, color=INK)
        ax.text(total + 1.2, y, f"{count}/{total} ({count / total:.0%})", ha="left", va="center", fontsize=11, color=INK)
    ax.set_xlim(-20, total + 18)
    ax.set_ylim(-0.8, rows - 0.2)
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)


def derive_tables(data: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    article = data["Article Evidence"]

    boundaries = (
        article["System_boundary_class"].value_counts().rename_axis("Category").reset_index(name="Studies")
    )
    methods = count_expanded(article, "Method_groups")
    databases = article["Database_group"].value_counts().rename_axis("Category").reset_index(name="Studies")
    robustness = pd.DataFrame(
        {
            "Analysis": ["Contribution analysis", "Sensitivity analysis", "Uncertainty analysis"],
            "Yes": [
                int(article["Contribution_analysis"].map(yes).sum()),
                int(article["Sensitivity_analysis"].map(yes).sum()),
                int(article["Uncertainty_analysis"].map(yes).sum()),
            ],
        }
    )
    robustness["No_or_not_reported"] = len(article) - robustness["Yes"]

    setting_scope = pd.crosstab(article["Scope_group"], article["Setting_group"])

    explicit_target = article.loc[
        ~article["Contaminant_groups"].map(lambda values: "No explicit remediation target" in values)
    ]
    technologies = sorted({item for values in explicit_target["Technology_groups"] for item in values})
    contaminants = sorted(
        {
            item
            for values in explicit_target["Contaminant_groups"]
            for item in values
            if item != "Other/unspecified"
        }
    )
    tech_contaminant = pd.DataFrame(0, index=contaminants, columns=technologies, dtype=int)
    for _, row in explicit_target.iterrows():
        for contaminant in row["Contaminant_groups"]:
            if contaminant not in tech_contaminant.index:
                continue
            for technology in row["Technology_groups"]:
                tech_contaminant.loc[contaminant, technology] += 1

    settings_subset = article.loc[article["Setting_group"].isin(["Marine/coastal", "Freshwater"])]
    setting_technologies = pd.DataFrame(
        0,
        index=sorted({item for values in settings_subset["Technology_groups"] for item in values}),
        columns=["Marine/coastal", "Freshwater"],
        dtype=int,
    )
    for _, row in settings_subset.iterrows():
        for technology in row["Technology_groups"]:
            setting_technologies.loc[technology, row["Setting_group"]] += 1

    outcomes = data["Outcome Evidence"].copy()
    signal_order = [
        "Higher-burden signal",
        "Hotspot-only signal",
        "Mixed/conditional signal",
        "Lower-burden signal",
    ]
    outcome_summary = (
        outcomes.groupby(["Impact_category", "Outcome_signal"])["Record_ID"]
        .nunique()
        .unstack(fill_value=0)
        .reindex(columns=signal_order, fill_value=0)
    )
    outcome_summary["Total_explicit_study_signals"] = outcome_summary.sum(axis=1)

    return {
        "boundaries": boundaries,
        "methods": methods,
        "databases": databases,
        "robustness": robustness,
        "setting_scope": setting_scope,
        "technology_contaminant": tech_contaminant,
        "setting_technologies": setting_technologies,
        "outcome_summary": outcome_summary,
    }


def export_tables(tables: dict[str, pd.DataFrame], output_dir: Path) -> None:
    table_dir = output_dir / "derived_tables"
    table_dir.mkdir(parents=True, exist_ok=True)
    for name, frame in tables.items():
        frame.to_csv(table_dir / f"{name}.csv", encoding="utf-8-sig")


def figure_2(article: pd.DataFrame, output_dir: Path, dpi: int) -> None:
    plot_data = article.dropna(subset=["Year"]).copy()
    plot_data["Year"] = plot_data["Year"].astype(int)
    year_scope = pd.crosstab(plot_data["Year"], plot_data["Scope_group"]).sort_index()
    desired = [
        "Remediation/risk control",
        "Management/placement",
        "Beneficial use/valorisation",
        "Resource recovery",
    ]
    columns = [column for column in desired if column in year_scope.columns]
    columns += [column for column in year_scope.columns if column not in columns]
    year_scope = year_scope.reindex(columns=columns, fill_value=0)
    palette = [NAVY, TEAL, GOLD, ORANGE, "#98A2B3"]
    fig, ax = plt.subplots(figsize=(9.2, 4.8))
    bottom = np.zeros(len(year_scope))
    for color, column in zip(palette, year_scope.columns):
        display_label = column.replace("valorisation", "valorization")
        ax.bar(year_scope.index, year_scope[column], bottom=bottom, label=display_label, color=color, width=0.8)
        bottom += year_scope[column].to_numpy()
    ax.set_ylabel("Number of studies")
    ax.set_xlabel("Publication year")
    ax.legend(frameon=False, ncol=2, fontsize=8)
    years = sorted(year_scope.index.unique())
    ax.set_xticks(years[::2] if len(years) > 10 else years)
    clean_axis(ax)
    save_figure(fig, output_dir, "Figure_02_Evidence_trend", dpi)


def figure_3(article: pd.DataFrame, tables: dict[str, pd.DataFrame], output_dir: Path, dpi: int) -> None:
    total = len(article)
    boundaries = tables["boundaries"]
    methods = tables["methods"]
    databases = tables["databases"]
    robustness = tables["robustness"]

    fig = plt.figure(figsize=(16.8, 12.8), constrained_layout=True)
    grid = fig.add_gridspec(3, 2, width_ratios=[1.10, 1], height_ratios=[1.10, 0.90, 0.98], wspace=0.16, hspace=0.12)

    ax = fig.add_subplot(grid[0, 0])
    order = [
        "Cradle-to-grave",
        "Cradle-to-gate",
        "Gate-to-grave",
        "Gate-to-gate",
        "Partial life cycle — material production and operation",
    ]
    counts = boundaries.set_index("Category")["Studies"].to_dict()
    stages = ["Upstream\nmaterials", "Treatment", "Downstream", "End-of-life"]
    stage_x = np.arange(4) * 1.55
    count_x = 6.30
    coverage = {
        "Cradle-to-grave": [1, 1, 1, 1],
        "Cradle-to-gate": [1, 1, 0, 0],
        "Gate-to-grave": [0, 1, 1, 1],
        "Gate-to-gate": [0, 1, 0, 0],
        "Partial life cycle — material production and operation": [0.5, 1, 0, 0],
    }
    available_order = [name for name in order if name in counts]
    for i, name in enumerate(available_order):
        y = len(available_order) - 1 - i
        ax.hlines(y, stage_x[0] - 0.05, stage_x[-1] + 0.05, color=LIGHT, lw=2, zorder=0)
        for x, value in enumerate(coverage[name]):
            if value == 1:
                ax.scatter(stage_x[x], y, s=290, color=TEAL, edgecolor="white", lw=1.2, zorder=2)
            elif value == 0.5:
                ax.scatter(stage_x[x], y, s=290, color=GOLD, edgecolor="white", lw=1.2, zorder=2)
            else:
                ax.scatter(stage_x[x], y, s=115, facecolor="white", edgecolor="#B7C0C8", lw=1.2, zorder=1)
        count = int(counts[name])
        ax.scatter(count_x, y, s=135 + 20 * count, color=NAVY, alpha=0.92, edgecolor="white", lw=1)
        ax.text(count_x, y, str(count), color="white", ha="center", va="center", fontweight="bold", fontsize=11)
    ax.set_yticks(np.arange(len(available_order)), [textwrap.fill(x, 24) for x in available_order[::-1]])
    ax.set_xticks(stage_x, stages)
    ax.text(count_x, len(available_order) - 0.35, "Studies", ha="center", va="bottom", fontsize=11, color=MID)
    ax.set_xlim(-0.80, 6.95)
    ax.set_ylim(-0.7, len(available_order) - 0.15)
    ax.spines[["left", "bottom"]].set_visible(False)
    ax.tick_params(length=0, labelsize=12)
    ax.legend(
        handles=[
            Line2D([0], [0], marker="o", color="none", markerfacecolor=TEAL, markeredgecolor="white", markersize=11, label="Included stage"),
            Line2D([0], [0], marker="o", color="none", markerfacecolor=GOLD, markeredgecolor="white", markersize=11, label="Selected/partial stage"),
            Line2D([0], [0], marker="o", color="none", markerfacecolor="white", markeredgecolor="#B7C0C8", markersize=9, label="Outside boundary"),
        ],
        frameon=False,
        ncol=3,
        loc="lower left",
        bbox_to_anchor=(0, -0.22),
        fontsize=11,
    )
    panel_label(ax, "A", fontsize=14)

    ax = fig.add_subplot(grid[0, 1])
    lollipop(ax, methods, total, BLUE)
    panel_label(ax, "B", fontsize=14)

    ax = fig.add_subplot(grid[1, 0])
    lollipop(ax, databases, total, ORANGE)
    panel_label(ax, "C", fontsize=14)

    ax = fig.add_subplot(grid[1, 1])
    normalization = article["Normalization_type"].value_counts().to_dict()
    weighting = article["Weighting_type"].value_counts().to_dict()
    labels = [
        "LCIA normalization / endpoint",
        "Comparative rescaling",
        "LCIA/method weighting",
        "Ad hoc weighting",
        "Decision-analysis weighting only",
    ]
    counts = [
        int(normalization.get("LCIA normalization / endpoint aggregation", 0)),
        int(normalization.get("Comparative rescaling to baseline or maximum", 0)),
        int(weighting.get("LCIA weighting / method-based single score", 0)),
        int(weighting.get("Ad hoc equal or study-specific aggregation", 0)),
        int(weighting.get("Decision-analysis weighting (not LCIA weighting)", 0)),
    ]
    dot_strip(ax, labels, counts, [BLUE, BLUE, ORANGE, ORANGE, GOLD], total)
    ax.axhline(2.5, color="#B7C0C8", lw=0.8)
    panel_label(ax, "D", fontsize=14)

    ax = fig.add_subplot(grid[2, 0])
    dot_strip(
        ax,
        robustness["Analysis"].tolist(),
        robustness["Yes"].astype(int).tolist(),
        [TEAL, BLUE, ORANGE],
        total,
    )
    panel_label(ax, "E", fontsize=14)

    ax = fig.add_subplot(grid[2, 1])
    score_columns = [
        "Score_FU_Boundaries",
        "Score_Foreground_Data",
        "Score_Background_DB",
        "Score_LCIA_Reporting",
        "Score_Allocation_Cutoff",
        "Score_Robustness",
    ]
    criterion_labels = [
        "FU and boundaries",
        "Foreground data",
        "Background database",
        "LCIA reporting",
        "Allocation and cut-off",
        "Interpretation robustness",
    ]
    score_counts = pd.DataFrame(
        {
            score: [int((pd.to_numeric(article[column], errors="coerce") == score).sum()) for column in score_columns]
            for score in [0, 1, 2]
        },
        index=criterion_labels,
    )
    shares = score_counts.div(total)
    y = np.arange(len(shares))
    left = np.zeros(len(shares))
    colors = ["#E8EDF1", "#A9C9DC", NAVY]
    for score, color in zip([0, 1, 2], colors):
        values = shares[score].to_numpy()
        bars = ax.barh(y, values, left=left, color=color, edgecolor="white", lw=0.8, label=f"Score {score}")
        for bar, count, share in zip(bars, score_counts[score], values):
            if count >= 4:
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_y() + bar.get_height() / 2,
                    str(int(count)),
                    ha="center",
                    va="center",
                    fontsize=11,
                    fontweight="bold",
                    color="white" if score == 2 else INK,
                )
        left += values
    ax.set_yticks(y, [textwrap.fill(label, 25) for label in criterion_labels])
    ax.invert_yaxis()
    ax.set_xlim(0, 1)
    ax.xaxis.set_major_formatter(lambda value, _: f"{value:.0%}")
    ax.set_xlabel("Share of studies")
    ax.tick_params(axis="both", labelsize=12)
    ax.xaxis.label.set_size(12)
    ax.grid(axis="x", color=LIGHT, lw=0.7, zorder=0)
    ax.legend(frameon=False, ncol=3, loc="lower center", bbox_to_anchor=(0.5, -0.27), fontsize=11)
    clean_axis(ax)
    panel_label(ax, "F", fontsize=14)
    save_figure(fig, output_dir, "Figure_04_Methodological_practices", dpi)


def figure_4(
    article: pd.DataFrame,
    tables: dict[str, pd.DataFrame],
    output_dir: Path,
    dpi: int,
) -> None:
    short = {
        "Biological/phytoremediation": "Bio / phyto",
        "Capping/in-situ amendment": "Capping / amendment",
        "Co-composting": "Co-composting",
        "Construction materials": "Construction",
        "Dredging/placement/disposal": "Dredging / disposal",
        "Land/horticultural use": "Land use",
        "Physical separation/flotation": "Physical separation",
        "Resource recovery": "Resource recovery",
        "Sediment bypass/coastal NBS": "Coastal NBS / bypass",
        "Solidification/stabilisation": "S/S",
        "Thermal/microwave/calcination": "Thermal",
        "Washing/leaching": "Washing / leaching",
    }

    technology_frequency = Counter()
    for technologies in article["Technology_groups"]:
        technology_frequency.update(set(technologies))
    evidence = tables["technology_contaminant"].T
    setting = tables["setting_technologies"].copy()
    technology_order = sorted(
        set(evidence.index) | set(setting.index),
        key=lambda technology: (-technology_frequency.get(technology, 0), technology),
    )
    contaminant_order = [
        "Metals/metalloids",
        "Organic contaminants",
        "Mixed/unspecified contamination",
        "Nutrients/phosphorus",
        "Microplastics",
        "Rare earth elements",
    ]
    evidence = evidence.reindex(index=technology_order, columns=contaminant_order, fill_value=0)
    setting = setting.reindex(index=technology_order, fill_value=0)

    fig = plt.figure(figsize=(17.0, 10.4), constrained_layout=True)
    grid = fig.add_gridspec(1, 2, width_ratios=[1.10, 1], wspace=0.10)

    ax = fig.add_subplot(grid[0, 0])
    array = evidence.to_numpy()
    image = ax.imshow(
        array,
        cmap=LinearSegmentedColormap.from_list("evidence", ["#F3F6F8", "#A9D7CF", TEAL, NAVY]),
        aspect="auto",
        vmin=0,
        vmax=max(1, array.max()),
    )
    contaminant_labels = {
        "Metals/metalloids": "Metals /\nmetalloids",
        "Organic contaminants": "Organic\ncontaminants",
        "Mixed/unspecified contamination": "Mixed /\nunspecified",
        "Nutrients/phosphorus": "Nutrients /\nphosphorus",
        "Microplastics": "Microplastics",
        "Rare earth elements": "Rare earth\nelements",
    }
    ax.set_xticks(np.arange(len(evidence.columns)), [contaminant_labels[x] for x in evidence.columns])
    ax.xaxis.tick_top()
    ax.tick_params(axis="x", pad=9, labelsize=12)
    ax.set_yticks(np.arange(len(evidence.index)), [short.get(x, x) for x in evidence.index])
    ax.tick_params(axis="y", labelsize=12)
    for i in range(array.shape[0]):
        for j in range(array.shape[1]):
            if array[i, j] > 0:
                ax.text(j, i, str(int(array[i, j])), ha="center", va="center", fontsize=11, color="white" if array[i, j] > 0.52 * array.max() else INK, fontweight="bold")
    ax.tick_params(length=0)
    ax.spines[:].set_visible(False)
    ax.set_xticks(np.arange(-0.5, len(evidence.columns), 1), minor=True)
    ax.set_yticks(np.arange(-0.5, len(evidence.index), 1), minor=True)
    ax.grid(which="minor", color="white", lw=1.2)
    ax.tick_params(which="minor", bottom=False, left=False)
    colorbar = fig.colorbar(image, ax=ax, orientation="horizontal", fraction=0.045, pad=0.045, aspect=35)
    explicit_target_n = int(
        (~article["Contaminant_groups"].map(lambda values: "No explicit remediation target" in values)).sum()
    )
    colorbar.set_label(
        f"Study count per pairing (explicit-target subset, n = {explicit_target_n})"
    )
    colorbar.set_ticks(np.arange(0, array.max() + 1, 1))
    colorbar.ax.tick_params(labelsize=12)
    colorbar.ax.xaxis.label.set_size(12)
    panel_label(ax, "A", fontsize=14)

    ax = fig.add_subplot(grid[0, 1])
    denominators = article["Setting_group"].value_counts()
    setting["Marine_share"] = setting["Marine/coastal"] / denominators["Marine/coastal"]
    setting["Fresh_share"] = setting["Freshwater"] / denominators["Freshwater"]
    setting["Max"] = setting[["Marine_share", "Fresh_share"]].max(axis=1)
    y = np.arange(len(setting))
    for yy, marine, fresh in zip(y, setting["Marine_share"], setting["Fresh_share"]):
        ax.hlines(yy, min(marine, fresh), max(marine, fresh), color="#B7C0C8", lw=1.8, zorder=1)
    ax.scatter(setting["Marine_share"], y, s=90, color=BLUE, edgecolor="white", lw=0.8, label=f"Marine/coastal (n = {denominators['Marine/coastal']})", zorder=3)
    ax.scatter(setting["Fresh_share"], y, s=90, color=TEAL, marker="D", edgecolor="white", lw=0.8, label=f"Freshwater (n = {denominators['Freshwater']})", zorder=3)
    ax.set_yticks(y, [])
    ax.set_ylim(len(setting) - 0.5, -0.5)
    upper_limit = max(0.62, setting["Max"].max() + 0.07)
    ax.set_xlim(-0.025, upper_limit)
    ax.set_xticks(np.arange(0, upper_limit + 0.001, 0.10))
    ax.xaxis.set_major_formatter(lambda value, _: f"{value:.0%}")
    ax.set_xlabel("Share of studies reporting each technology within setting")
    ax.tick_params(axis="x", labelsize=12)
    ax.xaxis.label.set_size(12)
    ax.grid(axis="both", color=LIGHT, lw=0.7)
    legend_handles, legend_labels = ax.get_legend_handles_labels()
    fig.legend(
        legend_handles,
        legend_labels,
        frameon=False,
        loc="center",
        bbox_to_anchor=(0.76, 0.082),
        ncol=2,
        fontsize=11,
    )
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#9AA4AD")
    panel_label(ax, "B", fontsize=14)
    save_figure(fig, output_dir, "Figure_03_Evidence_coverage", dpi)


def figure_5(
    article: pd.DataFrame,
    contribution: pd.DataFrame,
    outcomes: pd.DataFrame,
    tables: dict[str, pd.DataFrame],
    output_dir: Path,
    dpi: int,
) -> None:
    fig = plt.figure(figsize=(18.2, 9.8), constrained_layout=True)
    grid = fig.add_gridspec(1, 2, width_ratios=[0.96, 1.04], wspace=0.20)

    ax = fig.add_subplot(grid[0, 0])
    signal_order = [
        "Higher-burden signal",
        "Hotspot-only signal",
        "Mixed/conditional signal",
        "Lower-burden signal",
    ]
    signal_labels = ["Higher-burden\nsignal", "Hotspot-only\nsignal", "Mixed /\nconditional", "Lower-burden\nsignal"]
    signal_colors = [RED, ORANGE, GOLD, TEAL]
    summary = tables["outcome_summary"].copy()
    summary = summary.sort_values(["Total_explicit_study_signals", "Impact_category"] if "Impact_category" in summary.columns else ["Total_explicit_study_signals"], ascending=True)
    values = summary[signal_order].to_numpy()
    yy, xx = np.indices((len(summary), len(signal_order)))
    mask = values.ravel() > 0
    ax.scatter(xx.ravel()[mask], yy.ravel()[mask], s=80 + 62 * values.ravel()[mask], c=np.array(signal_colors)[xx.ravel()[mask]], edgecolor="white", lw=1.0, zorder=3)
    for i in range(values.shape[0]):
        for j in range(values.shape[1]):
            if values[i, j] > 0:
                ax.text(j, i, str(int(values[i, j])), ha="center", va="center", color="white" if values[i, j] >= 4 else INK, fontsize=10.8, fontweight="bold", zorder=4)
        ax.text(3.62, i, f"n = {int(summary.iloc[i]['Total_explicit_study_signals'])}", va="center", ha="left", fontsize=10.2, color=MID)
    ax.set_xticks(range(4), signal_labels)
    ax.set_yticks(range(len(summary)), [textwrap.fill(str(x), 24) for x in summary.index])
    ax.set_xlim(-0.55, 4.18)
    ax.set_ylim(-0.55, len(summary) - 0.45)
    ax.tick_params(length=0, labelsize=11.2)
    ax.set_xticks(np.arange(-0.5, 4, 1), minor=True)
    ax.set_yticks(np.arange(-0.5, len(summary), 1), minor=True)
    ax.grid(which="minor", color=LIGHT, lw=0.65)
    ax.tick_params(which="minor", bottom=False, left=False)
    ax.spines[:].set_visible(False)
    panel_label(ax, "A", fontsize=14)

    ax = fig.add_subplot(grid[0, 1])
    technology_lookup = article.set_index("Record_ID")["Technology_groups"].to_dict()
    cooccurrences: list[tuple[str, str]] = []
    for _, row in contribution.iterrows():
        groups = [item.strip() for item in str(row["Broad_process_groups"]).split(";") if item.strip()]
        for group in groups:
            for technology in technology_lookup.get(row["Record_ID"], []):
                cooccurrences.append((group, technology))
    cooccurrence = pd.DataFrame(cooccurrences, columns=["Contributor", "Technology"])
    top_technologies = cooccurrence["Technology"].value_counts().head(9).index.tolist()
    top_groups = cooccurrence["Contributor"].value_counts().index.tolist()
    matrix = pd.crosstab(cooccurrence["Contributor"], cooccurrence["Technology"]).reindex(index=top_groups, columns=top_technologies, fill_value=0)
    array = matrix.to_numpy()
    yy, xx = np.indices(array.shape)
    mask = array.ravel() > 0
    ax.scatter(
        xx.ravel()[mask],
        yy.ravel()[mask],
        s=60 + 130 * array.ravel()[mask],
        c=array.ravel()[mask],
        cmap=LinearSegmentedColormap.from_list("bubbles", ["#B8DCD5", TEAL, NAVY]),
        vmin=1,
        vmax=max(2, array.max()),
        edgecolor="white",
        lw=0.8,
    )
    for i in range(array.shape[0]):
        for j in range(array.shape[1]):
            if array[i, j] > 0:
                ax.text(j, i, str(array[i, j]), ha="center", va="center", color="white" if array[i, j] >= max(3, array.max() * 0.55) else INK, fontsize=10.8, fontweight="bold")
    technology_labels = {
        "Construction materials": "Construction",
        "Dredging/placement/disposal": "Dredging / disposal",
        "Land/horticultural use": "Land use",
        "Capping/in-situ amendment": "Capping / amendment",
        "Solidification/stabilisation": "S/S",
        "Washing/leaching": "Washing / leaching",
        "Biological/phytoremediation": "Bio / phyto",
        "Sediment bypass/coastal NBS": "Coastal NBS",
    }
    ax.set_xticks(range(len(top_technologies)), [technology_labels.get(x, x) for x in top_technologies], rotation=42, ha="right")
    ax.set_yticks(range(len(top_groups)), [textwrap.fill(x, 23) for x in top_groups])
    ax.set_xlim(-0.6, len(top_technologies) - 0.4)
    ax.set_ylim(len(top_groups) - 0.5, -0.5)
    ax.tick_params(length=0, labelsize=10.8)
    ax.spines[:].set_visible(False)
    ax.set_xticks(np.arange(-0.5, len(top_technologies), 1), minor=True)
    ax.set_yticks(np.arange(-0.5, len(top_groups), 1), minor=True)
    ax.grid(which="minor", color=LIGHT, lw=0.7)
    ax.tick_params(which="minor", bottom=False, left=False)
    panel_label(ax, "B", fontsize=14)

    save_figure(fig, output_dir, "Figure_05_Environmental_outcomes_contributors", dpi)


def figure_6(article: pd.DataFrame, output_dir: Path, dpi: int) -> None:
    circularity_columns = [
        "Circularity_waste_prevention",
        "Circularity_sediment_recycling_recovery",
        "Circularity_byproduct_or_secondary_input",
        "Circularity_avoided_burden_modelled",
    ]
    short_names = ["Waste prevention", "Sediment recycling / recovery", "Secondary input", "Avoided-burden modeling"]
    patterns = Counter(tuple(yes(row[column]) for column in circularity_columns) for _, row in article.iterrows())
    top = patterns.most_common(9)
    x = np.arange(len(top))
    values = np.array([value for _, value in top])

    fig, ax = plt.subplots(figsize=(11.8, 7.8), constrained_layout=False)
    ax.bar(x, values, color=NAVY, width=0.67)
    for xx, value in zip(x, values):
        ax.text(xx, value + 0.35, str(value), ha="center", va="bottom", color=INK, fontsize=11.5, fontweight="bold")
    base = -2.2
    step = 1.05
    for index, (pattern, _) in enumerate(top):
        active = [i for i, present in enumerate(pattern) if present]
        if len(active) > 1:
            positions = [base - i * step for i in active]
            ax.vlines(index, min(positions), max(positions), color=INK, lw=1.2)
        for i, present in enumerate(pattern):
            ax.scatter(index, base - i * step, s=82 if present else 48, color=TEAL if present else LIGHT, edgecolor="white" if present else "none", lw=0.5, zorder=3)
    for i, label in enumerate(short_names):
        ax.text(
            -0.035,
            base - i * step,
            label,
            transform=ax.get_yaxis_transform(),
            ha="right",
            va="center",
            fontsize=11.5,
            color=INK,
            clip_on=False,
        )
    ax.set_xlim(-0.50, len(top) - 0.25)
    ax.set_ylim(base - (len(short_names) - 1) * step - 0.72, max(values) + 2.0)
    ax.set_xticks(x, [f"I{i + 1}" for i in range(len(top))])
    ax.set_yticks(np.arange(0, max(values) + 1, 5))
    ax.set_ylabel("Studies")
    ax.set_xlabel(f"Study-level combination of circularity elements (n = {len(article)})")
    ax.tick_params(axis="both", labelsize=11.2)
    ax.xaxis.label.set_size(11.8)
    ax.yaxis.label.set_size(11.8)
    ax.grid(axis="y", color=LIGHT, lw=0.8, zorder=0)
    clean_axis(ax)
    fig.subplots_adjust(left=0.31, right=0.98, bottom=0.14, top=0.96)
    save_figure(fig, output_dir, "Figure_06_Circularity_modeling", dpi)


def categorical_distance(frame: pd.DataFrame, base_fields: list[str], method_fields: list[str], ignore_missing: bool) -> np.ndarray:
    n = len(frame)
    distance = np.zeros((n, n), dtype=float)
    for i in range(n):
        for j in range(i + 1, n):
            comparisons: list[bool] = []
            for column in base_fields:
                first, second = str(frame.iloc[i][column]), str(frame.iloc[j][column])
                if ignore_missing and ("not reported" in first.lower() or "not reported" in second.lower()):
                    continue
                comparisons.append(first != second)
            comparisons.extend(frame.iloc[i][column] != frame.iloc[j][column] for column in method_fields)
            distance[i, j] = distance[j, i] = sum(comparisons) / len(comparisons) if comparisons else 0.0
    return distance


def jaccard_distance(frame: pd.DataFrame, base_fields: list[str], method_fields: list[str]) -> np.ndarray:
    feature_sets: list[set[str]] = []
    for _, row in frame.iterrows():
        features = {f"{column}={row[column]}" for column in base_fields}
        features.update(column for column in method_fields if row[column] == "Yes")
        feature_sets.append(features)
    n = len(frame)
    distance = np.zeros((n, n), dtype=float)
    for i in range(n):
        for j in range(i + 1, n):
            union = feature_sets[i] | feature_sets[j]
            intersection = feature_sets[i] & feature_sets[j]
            distance[i, j] = distance[j, i] = 1 - len(intersection) / len(union) if union else 0.0
    return distance


def average_linkage(distance: np.ndarray) -> tuple[list[tuple[int, int, float, int, int]], dict[int, tuple[int | None, int | None, float, list[int]]], int]:
    n = distance.shape[0]
    clusters = {i: [i] for i in range(n)}
    nodes: dict[int, tuple[int | None, int | None, float, list[int]]] = {i: (None, None, 0.0, [i]) for i in range(n)}
    next_id = n
    merges: list[tuple[int, int, float, int, int]] = []
    while len(clusters) > 1:
        keys = sorted(clusters)
        best: tuple[tuple[float, int, int], int, int, float] | None = None
        for position, first in enumerate(keys):
            for second in keys[position + 1 :]:
                value = float(np.mean([distance[i, j] for i in clusters[first] for j in clusters[second]]))
                tie = (value, min(clusters[first]), min(clusters[second]))
                if best is None or tie < best[0]:
                    best = (tie, first, second, value)
        assert best is not None
        _, first, second, value = best
        members = clusters.pop(first) + clusters.pop(second)
        clusters[next_id] = members
        nodes[next_id] = (first, second, value, members)
        merges.append((first, second, value, len(members), next_id))
        next_id += 1
    return merges, nodes, next_id - 1


def cut_clusters(merges: list[tuple[int, int, float, int, int]], n: int, k: int = 3) -> np.ndarray:
    clusters: dict[int, set[int]] = {i: {i} for i in range(n)}
    for first, second, _, _, new in merges:
        clusters[new] = clusters.pop(first) | clusters.pop(second)
        if len(clusters) == k:
            break
    ordered = sorted(clusters.values(), key=min)
    labels = np.zeros(n, dtype=int)
    for label, members in enumerate(ordered, 1):
        for member in members:
            labels[member] = label
    return labels


def adjusted_rand(first: np.ndarray, second: np.ndarray) -> float:
    def combinations_two(value: int) -> int:
        return value * (value - 1) // 2

    table = pd.crosstab(pd.Series(first), pd.Series(second)).to_numpy()
    sum_combined = sum(combinations_two(int(value)) for value in table.ravel())
    row_combined = sum(combinations_two(int(value)) for value in table.sum(axis=1))
    column_combined = sum(combinations_two(int(value)) for value in table.sum(axis=0))
    total = combinations_two(len(first))
    expected = row_combined * column_combined / total if total else 0
    maximum = (row_combined + column_combined) / 2
    return (sum_combined - expected) / (maximum - expected) if maximum != expected else 1.0


def figure_s1(article: pd.DataFrame, output_dir: Path, dpi: int) -> pd.DataFrame:
    base_fields = [
        "System_boundary_class",
        "Database_group",
        "Normalization_type",
        "Weighting_type",
        "Contribution_analysis",
        "Sensitivity_analysis",
        "Uncertainty_analysis",
    ]
    clustered = article.set_index("Record_ID")[base_fields].astype(str).copy()
    methods = sorted({method for values in article["Method_groups"] for method in values})
    for method in methods:
        clustered[f"Method::{method}"] = article.set_index("Record_ID")["Method_groups"].map(lambda values: "Yes" if method in values else "No")
    method_fields = [column for column in clustered.columns if column.startswith("Method::")]
    ids = clustered.index.tolist()
    n = len(ids)

    hamming = categorical_distance(clustered, base_fields, method_fields, ignore_missing=False)
    missing_excluded = categorical_distance(clustered, base_fields, method_fields, ignore_missing=True)
    jaccard = jaccard_distance(clustered, base_fields, method_fields)
    results = {}
    for name, distance in [("Hamming", hamming), ("Missing-excluded", missing_excluded), ("Jaccard", jaccard)]:
        merges, nodes, root = average_linkage(distance)
        results[name] = (distance, merges, nodes, root, cut_clusters(merges, n, 3))

    stability = pd.DataFrame(
        [
            ["Hamming vs missing-excluded", adjusted_rand(results["Hamming"][4], results["Missing-excluded"][4])],
            ["Hamming vs Jaccard", adjusted_rand(results["Hamming"][4], results["Jaccard"][4])],
            ["Missing-excluded vs Jaccard", adjusted_rand(results["Missing-excluded"][4], results["Jaccard"][4])],
        ],
        columns=["Comparison", "Adjusted_Rand_Index"],
    )

    _, _, nodes, root, _ = results["Hamming"]

    def leaf_order(node: int) -> list[int]:
        left, right, _, _ = nodes[node]
        if left is None or right is None:
            return [node]
        left_order = leaf_order(left)
        right_order = leaf_order(right)
        return left_order + right_order if min(left_order) <= min(right_order) else right_order + left_order

    order = leaf_order(root)
    positions = {leaf: position for position, leaf in enumerate(order)}
    coordinates: dict[int, tuple[float, float]] = {}

    fig = plt.figure(figsize=(15, 8.8), constrained_layout=True)
    grid = fig.add_gridspec(1, 2, width_ratios=[2.6, 1], wspace=0.08)
    ax = fig.add_subplot(grid[0, 0])

    def draw(node: int) -> tuple[float, float]:
        left, right, height, _ = nodes[node]
        if left is None or right is None:
            coordinates[node] = (0.0, float(positions[node]))
            return coordinates[node]
        left_x, left_y = draw(left)
        right_x, right_y = draw(right)
        y = (left_y + right_y) / 2
        ax.plot([left_x, height], [left_y, left_y], color=NAVY, lw=1)
        ax.plot([right_x, height], [right_y, right_y], color=NAVY, lw=1)
        ax.plot([height, height], [left_y, right_y], color=NAVY, lw=1)
        coordinates[node] = (height, y)
        return coordinates[node]

    draw(root)
    ax.set_yticks(range(n), [ids[index] for index in order], fontsize=7.2)
    ax.set_xlabel("Average methodological dissimilarity")
    ax.invert_yaxis()
    clean_axis(ax)
    panel_label(ax, "A")

    ax = fig.add_subplot(grid[0, 1])
    labels = ["Hamming", "Missing-excluded", "Jaccard"]
    agreement = np.full((3, 3), np.nan)
    locations = {
        "Hamming vs missing-excluded": (0, 1),
        "Hamming vs Jaccard": (0, 2),
        "Missing-excluded vs Jaccard": (1, 2),
    }
    for _, row in stability.iterrows():
        i, j = locations[row["Comparison"]]
        agreement[i, j] = agreement[j, i] = row["Adjusted_Rand_Index"]
    image = ax.imshow(np.ma.masked_invalid(agreement), cmap=LinearSegmentedColormap.from_list("ari", [RED, "#F7F7F5", TEAL]), vmin=-0.05, vmax=0.15)
    for i in range(3):
        for j in range(3):
            if np.isnan(agreement[i, j]):
                ax.text(j, i, "—", ha="center", va="center", color=MID)
            else:
                ax.text(j, i, f"{agreement[i, j]:.3f}", ha="center", va="center", color=INK, fontweight="bold")
    ax.set_xticks(range(3), labels, rotation=35, ha="right")
    ax.set_yticks(range(3), labels)
    ax.tick_params(length=0)
    ax.spines[:].set_visible(False)
    colorbar = fig.colorbar(image, ax=ax, fraction=0.05, pad=0.04)
    colorbar.set_label("Adjusted Rand index")
    ax.text(0.5, -0.24, "Low and inconsistent agreement indicates that the three-cluster solution is not stable.", transform=ax.transAxes, ha="center", va="top", color=MID, fontsize=9, wrap=True)
    panel_label(ax, "B")
    save_figure(fig, output_dir, "Figure_S1_Clustering_audit", dpi)
    return stability


def write_manifest(output_dir: Path, workbook: Path, article_count: int, figures: list[str]) -> None:
    rows = []
    for stem in figures:
        rows.append(
            {
                "Figure": stem,
                "PNG": f"{stem}.png",
                "PDF": f"{stem}.pdf",
                "Source_workbook": workbook.name,
                "Included_studies": article_count,
            }
        )
    pd.DataFrame(rows).to_csv(output_dir / "figure_manifest.csv", index=False, encoding="utf-8-sig")


def main() -> None:
    args = parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    data = load_workbook(args.workbook)
    article = data["Article Evidence"]
    tables = derive_tables(data)
    export_tables(tables, args.output_dir)

    figure_2(article, args.output_dir, args.dpi)
    figure_3(article, tables, args.output_dir, args.dpi)
    figure_4(article, tables, args.output_dir, args.dpi)
    figure_5(article, data["Contribution Analysis"], data["Outcome Evidence"], tables, args.output_dir, args.dpi)
    figure_6(article, args.output_dir, args.dpi)
    stability = figure_s1(article, args.output_dir, args.dpi)
    stability.to_csv(args.output_dir / "derived_tables" / "cluster_metric_stability.csv", index=False, encoding="utf-8-sig")

    stems = [
        "Figure_02_Evidence_trend",
        "Figure_03_Evidence_coverage",
        "Figure_04_Methodological_practices",
        "Figure_05_Environmental_outcomes_contributors",
        "Figure_06_Circularity_modeling",
        "Figure_S1_Clustering_audit",
    ]
    write_manifest(args.output_dir, args.workbook, len(article), stems)
    print(f"Generated {len(stems)} figures from {args.workbook}")
    print(f"Output directory: {args.output_dir.resolve()}")


if __name__ == "__main__":
    main()
