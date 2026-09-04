from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
import yaml

from .config import ASSUMPTIONS_PATH


@dataclass(frozen=True)
class ModelBundle:
    assumptions: dict[str, Any]
    financials: pd.DataFrame
    scenarios: dict[str, pd.DataFrame]
    market: pd.DataFrame
    use_of_proceeds: pd.DataFrame
    unit_economics: pd.DataFrame
    cohorts: pd.DataFrame
    transaction: dict[str, float]


def load_yaml(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        payload = yaml.safe_load(handle)
    if not isinstance(payload, dict):
        raise ValueError(f"{path} must contain a mapping at the root")
    return payload


def _require_equal_lengths(name: str, expected: int, values: list[Any]) -> None:
    if len(values) != expected:
        raise ValueError(
            f"{name} has {len(values)} values; expected {expected}"
        )


def _build_financials(assumptions: dict[str, Any]) -> pd.DataFrame:
    years = assumptions["years"]
    frame = pd.DataFrame(index=pd.Index(years, name="year"))
    financials = assumptions["financials"]

    for segment, values in financials["revenue_by_segment"].items():
        _require_equal_lengths(f"revenue_by_segment.{segment}", len(years), values)
        frame[f"revenue_{segment}"] = values

    frame["revenue"] = frame.filter(like="revenue_").sum(axis=1)
    frame["revenue_growth"] = frame["revenue"].pct_change()

    direct_columns = {
        "gross_margin": "gross_margin",
        "research_and_development": "r_and_d",
        "sales_and_marketing": "sales_and_marketing",
        "general_and_administrative": "g_and_a",
        "capital_expenditures": "capex",
        "working_capital_change": "working_capital_change",
        "stock_compensation": "stock_compensation",
    }
    for source_name, column_name in direct_columns.items():
        values = financials[source_name]
        _require_equal_lengths(source_name, len(years), values)
        frame[column_name] = values

    frame["cost_of_revenue"] = frame["revenue"] * (1 - frame["gross_margin"])
    frame["gross_profit"] = frame["revenue"] - frame["cost_of_revenue"]
    frame["operating_expenses"] = (
        frame["r_and_d"] + frame["sales_and_marketing"] + frame["g_and_a"]
    )
    frame["adjusted_operating_income"] = (
        frame["gross_profit"] - frame["operating_expenses"]
    )
    frame["adjusted_operating_margin"] = (
        frame["adjusted_operating_income"] / frame["revenue"]
    )
    frame["free_cash_flow"] = (
        frame["adjusted_operating_income"]
        + frame["stock_compensation"]
        - frame["capex"]
        - frame["working_capital_change"]
    )
    frame["free_cash_flow_margin"] = frame["free_cash_flow"] / frame["revenue"]

    for metric, values in assumptions["kpis"].items():
        _require_equal_lengths(f"kpis.{metric}", len(years), values)
        frame[metric] = values

    frame["paid_conversion"] = (
        frame["paid_individuals_m"] / frame["monthly_active_builders_m"]
    )
    frame["individual_arpu"] = (
        frame["revenue_individual"] / frame["paid_individuals_m"]
    )
    frame["enterprise_revenue_per_customer_k"] = (
        frame["revenue_enterprise"] * 1000 / frame["enterprise_customers"]
    )
    return frame


def _build_scenarios(
    assumptions: dict[str, Any], financials: pd.DataFrame
) -> dict[str, pd.DataFrame]:
    scenario_years = list(range(2026, 2032))
    base = financials.loc[scenario_years]
    scenarios: dict[str, pd.DataFrame] = {}

    for name, scenario in assumptions["scenarios"].items():
        for key in (
            "revenue_multiplier",
            "gross_margin_delta",
            "opex_multiplier",
        ):
            _require_equal_lengths(
                f"scenarios.{name}.{key}", len(scenario_years), scenario[key]
            )

        frame = pd.DataFrame(index=pd.Index(scenario_years, name="year"))
        frame["revenue"] = (
            base["revenue"].to_numpy() * scenario["revenue_multiplier"]
        )
        frame["revenue_growth"] = frame["revenue"].pct_change()
        frame["gross_margin"] = (
            base["gross_margin"].to_numpy() + scenario["gross_margin_delta"]
        )
        frame["gross_profit"] = frame["revenue"] * frame["gross_margin"]
        frame["operating_expenses"] = (
            base["operating_expenses"].to_numpy() * scenario["opex_multiplier"]
        )
        frame["adjusted_operating_income"] = (
            frame["gross_profit"] - frame["operating_expenses"]
        )
        frame["adjusted_operating_margin"] = (
            frame["adjusted_operating_income"] / frame["revenue"]
        )
        frame["free_cash_flow"] = (
            frame["adjusted_operating_income"]
            + base["stock_compensation"].to_numpy()
            - base["capex"].to_numpy()
            - base["working_capital_change"].to_numpy()
        )
        frame["free_cash_flow_margin"] = (
            frame["free_cash_flow"] / frame["revenue"]
        )
        scenarios[name] = frame
    return scenarios


def _build_market(assumptions: dict[str, Any]) -> pd.DataFrame:
    frame = pd.DataFrame(assumptions["market"]["segments"])
    frame["tam"] = frame["population_m"] * frame["annual_revenue_per_unit"]
    frame["tam_share"] = frame["tam"] / frame["tam"].sum()
    frame["sam_share"] = frame["sam"] / frame["sam"].sum()
    return frame


def _build_use_of_proceeds(assumptions: dict[str, Any]) -> pd.DataFrame:
    frame = pd.DataFrame(assumptions["use_of_proceeds"])
    primary = assumptions["transaction"]["primary"]
    if round(frame["amount"].sum(), 6) != round(primary, 6):
        raise ValueError("Use-of-proceeds amounts must sum to primary proceeds")
    if round(frame["percent"].sum(), 6) != 1:
        raise ValueError("Use-of-proceeds percentages must sum to 100%")
    return frame


def _build_unit_economics(assumptions: dict[str, Any]) -> pd.DataFrame:
    payload = assumptions["unit_economics"]
    years = payload["years"]
    frame = pd.DataFrame(index=pd.Index(years, name="year"))
    for key, values in payload.items():
        if key == "years":
            continue
        _require_equal_lengths(f"unit_economics.{key}", len(years), values)
        frame[key] = values
    return frame


def _build_cohorts(assumptions: dict[str, Any]) -> pd.DataFrame:
    payload = assumptions["cohorts"]
    years = payload["years"]
    records: list[dict[str, Any]] = []
    months = {
        "month_12_multiple": 12,
        "month_24_multiple": 24,
        "month_36_multiple": 36,
    }
    for position, cohort_year in enumerate(years):
        starting_arr = payload["starting_arr"][position]
        records.append(
            {
                "cohort": cohort_year,
                "month": 0,
                "arr": starting_arr,
                "multiple": 1.0,
            }
        )
        for field, month in months.items():
            multiple = payload[field][position]
            if multiple is None:
                continue
            records.append(
                {
                    "cohort": cohort_year,
                    "month": month,
                    "arr": starting_arr * multiple,
                    "multiple": multiple,
                }
            )
    return pd.DataFrame(records)


def _build_transaction(assumptions: dict[str, Any]) -> dict[str, float]:
    source = assumptions["transaction"]
    post_money = source["pre_money"] + source["primary"]
    return {
        **source,
        "post_money": post_money,
        "primary_ownership": source["primary"] / post_money,
        "secondary_ownership": source["secondary"] / post_money,
        "total_buyer_ownership": source["total_raise"] / post_money,
        "valuation_step_up": source["pre_money"] / source["prior_round_valuation"],
    }


def build_model(path: Path = ASSUMPTIONS_PATH) -> ModelBundle:
    assumptions = load_yaml(path)
    financials = _build_financials(assumptions)
    return ModelBundle(
        assumptions=assumptions,
        financials=financials,
        scenarios=_build_scenarios(assumptions, financials),
        market=_build_market(assumptions),
        use_of_proceeds=_build_use_of_proceeds(assumptions),
        unit_economics=_build_unit_economics(assumptions),
        cohorts=_build_cohorts(assumptions),
        transaction=_build_transaction(assumptions),
    )
