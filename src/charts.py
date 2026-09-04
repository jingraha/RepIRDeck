from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np

from .config import COLORS
from .model import ModelBundle


def _hex(name: str) -> str:
    return f"#{COLORS[name]}"


def _base_axes(ax) -> None:
    ax.set_facecolor(_hex("cream"))
    ax.figure.set_facecolor(_hex("cream"))
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.spines["bottom"].set_color(_hex("gray_300"))
    ax.grid(axis="y", color=_hex("gray_200"), linewidth=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(colors=_hex("gray_700"), labelsize=8, length=0)


def render_workbook_charts(model: ModelBundle, output_dir: Path) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    results: dict[str, Path] = {}

    years = np.arange(2026, 2032)
    individual = model.financials.loc[years, "revenue_individual"].to_numpy()
    enterprise = model.financials.loc[years, "revenue_enterprise"].to_numpy()
    usage = model.financials.loc[years, "revenue_usage_and_deploy"].to_numpy()
    fig, ax = plt.subplots(figsize=(8.8, 4.3))
    _base_axes(ax)
    ax.bar(years, individual, color=_hex("orange"), label="Individual")
    ax.bar(
        years,
        enterprise,
        bottom=individual,
        color=_hex("orange_mid"),
        label="Enterprise",
    )
    ax.bar(
        years,
        usage,
        bottom=individual + enterprise,
        color=_hex("coral"),
        label="Agent use and hosting",
    )
    totals = individual + enterprise + usage
    for year, total in zip(years, totals, strict=True):
        ax.text(
            year,
            total + totals.max() * 0.025,
            f"${total / 1000:.2f}B" if total >= 1000 else f"${total:.0f}M",
            ha="center",
            va="bottom",
            fontsize=8,
            fontweight="bold",
            color=_hex("ink"),
        )
    ax.set_title("Revenue by source", loc="left", fontsize=15, fontweight="bold")
    ax.set_ylabel("$M", fontsize=9)
    ax.set_xticks(years, [f"{year}E" for year in years])
    ax.legend(frameon=False, ncol=3, loc="upper left")
    fig.tight_layout()
    path = output_dir / "revenue_by_source.png"
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)
    results["revenue"] = path

    fig, ax = plt.subplots(figsize=(8.8, 3.6))
    _base_axes(ax)
    left = 0.0
    for _, item in model.use_of_proceeds.iterrows():
        ax.barh(
            ["Primary capital"],
            [item["amount"]],
            left=[left],
            color=_hex(item["color"]),
            label=f"{item['category']}  ${item['amount']:.0f}M",
        )
        if item["percent"] >= 0.10:
            ax.text(
                left + item["amount"] / 2,
                0,
                f"{item['percent']:.0%}",
                ha="center",
                va="center",
                fontsize=9,
                fontweight="bold",
                color="white",
            )
        left += item["amount"]
    ax.set_title("Use of $800M primary capital", loc="left", fontsize=15, fontweight="bold")
    ax.set_xlabel("$M", fontsize=9)
    ax.set_xlim(0, model.transaction["primary"])
    ax.legend(frameon=False, ncol=2, loc="upper center", bbox_to_anchor=(0.5, -0.18))
    fig.tight_layout()
    path = output_dir / "use_of_proceeds.png"
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)
    results["use_of_proceeds"] = path

    cash = model.monthly_cash_flow.loc[: model.cash_summary["ipo_target_date"]]
    no_series_e = cash["ending_cash_without_series_e"].where(
        cash.index <= model.cash_summary["funding_need_date"]
    )
    fig, ax = plt.subplots(figsize=(11.8, 4.2))
    _base_axes(ax)
    ax.plot(
        cash.index,
        cash["ending_cash"],
        color=_hex("orange"),
        linewidth=2.8,
        label="Base with Series E",
    )
    ax.plot(
        cash.index,
        cash["ending_cash_downside"],
        color=_hex("coral"),
        linewidth=2.4,
        label="Downside with Series E",
    )
    ax.plot(
        cash.index,
        no_series_e,
        color=_hex("gray_500"),
        linewidth=2.4,
        label="No Series E",
    )
    ax.plot(
        cash.index,
        cash["minimum_cash"],
        color=_hex("orange_dark"),
        linewidth=1.5,
        linestyle="--",
        label="Minimum cash",
    )
    ax.axvline(
        model.cash_summary["financing_close_date"],
        color=_hex("orange_dark"),
        linewidth=1,
        linestyle=":",
    )
    ax.axvline(
        model.cash_summary["ipo_target_date"],
        color=_hex("orange"),
        linewidth=1,
        linestyle=":",
    )
    ax.set_title(
        "Monthly ending cash through the IPO target",
        loc="left",
        fontsize=15,
        fontweight="bold",
    )
    ax.set_ylabel("$M", fontsize=9)
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=6))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b-%y"))
    ax.legend(frameon=False, ncol=4, loc="upper left")
    fig.autofmt_xdate(rotation=0)
    fig.tight_layout()
    path = output_dir / "monthly_cash.png"
    fig.savefig(path, dpi=180, bbox_inches="tight")
    plt.close(fig)
    results["cash"] = path
    return results
