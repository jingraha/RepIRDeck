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
    monthly_cash_flow: pd.DataFrame
    cash_summary: dict[str, Any]
    income_statement: pd.DataFrame
    cash_flow_statement: pd.DataFrame
    balance_sheet: pd.DataFrame


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


def _build_monthly_cash_flow(
    assumptions: dict[str, Any],
    financials: pd.DataFrame,
    scenarios: dict[str, pd.DataFrame],
    transaction: dict[str, float],
) -> tuple[pd.DataFrame, dict[str, Any]]:
    cash = assumptions["cash"]
    dates = pd.date_range(
        cash["monthly_start"],
        cash["monthly_end"],
        freq="ME",
    )
    revenue_weights = pd.Series(
        [1 / 12] * 12,
        index=range(1, 13),
        dtype=float,
    )

    records: list[dict[str, Any]] = []
    base_cash = float(cash["beginning_cash"])
    no_financing_cash = float(cash["beginning_cash"])
    downside_cash = float(cash["beginning_cash"])
    financing_close = pd.Timestamp(cash["financing_close_date"])

    for date in dates:
        annual = financials.loc[date.year]
        weight = float(revenue_weights.loc[date.month])
        revenue = annual["revenue"] * weight
        gross_profit = revenue * annual["gross_margin"]
        operating_expenses = annual["operating_expenses"] / 12
        adjusted_operating_income = gross_profit - operating_expenses
        stock_compensation = annual["stock_compensation"] / 12
        capex = annual["capex"] / 12
        working_capital_change = annual["working_capital_change"] * weight
        free_cash_flow = (
            adjusted_operating_income
            + stock_compensation
            - capex
            - working_capital_change
        )

        downside_annual = scenarios["downside"].loc[date.year]
        downside_revenue = downside_annual["revenue"] * weight
        downside_gross_profit = downside_revenue * downside_annual["gross_margin"]
        downside_operating_expenses = downside_annual["operating_expenses"] / 12
        downside_free_cash_flow = (
            downside_gross_profit
            - downside_operating_expenses
            + stock_compensation
            - capex
            - working_capital_change
        )

        primary_financing = (
            transaction["primary"] if date == financing_close else 0.0
        )
        financing_fees = (
            float(cash["financing_fees"]) if date == financing_close else 0.0
        )

        beginning_cash = base_cash
        base_cash += free_cash_flow + primary_financing - financing_fees
        no_financing_cash += free_cash_flow
        downside_cash += (
            downside_free_cash_flow + primary_financing - financing_fees
        )

        records.append(
            {
                "date": date,
                "beginning_cash": beginning_cash,
                "revenue": revenue,
                "gross_profit": gross_profit,
                "operating_expenses": operating_expenses,
                "adjusted_operating_income": adjusted_operating_income,
                "stock_compensation": stock_compensation,
                "capital_expenditures": capex,
                "working_capital_change": working_capital_change,
                "free_cash_flow": free_cash_flow,
                "primary_financing": primary_financing,
                "financing_fees": financing_fees,
                "ending_cash": base_cash,
                "ending_cash_without_series_e": no_financing_cash,
                "ending_cash_downside": downside_cash,
                "minimum_cash": float(cash["minimum_cash"]),
            }
        )

    frame = pd.DataFrame(records).set_index("date")
    below_minimum = frame[
        frame["ending_cash_without_series_e"] < float(cash["minimum_cash"])
    ]
    funding_need_date = below_minimum.index[0] if not below_minimum.empty else None
    ipo_target = pd.Timestamp(cash["ipo_target_date"])
    frame["milestone"] = ""
    frame.loc[financing_close, "milestone"] = "Series E close"
    if funding_need_date is not None:
        existing = frame.loc[funding_need_date, "milestone"]
        note = "No-Series-E cash below minimum"
        frame.loc[funding_need_date, "milestone"] = (
            f"{existing} | {note}" if existing else note
        )
    frame.loc[ipo_target, "milestone"] = "IPO target"
    ipo_row = frame.loc[frame.index <= ipo_target].iloc[-1]
    summary = {
        "as_of_date": pd.Timestamp(cash["as_of_date"]),
        "beginning_cash": float(cash["beginning_cash"]),
        "minimum_cash": float(cash["minimum_cash"]),
        "funding_need_date": funding_need_date,
        "financing_close_date": financing_close,
        "net_primary_proceeds": transaction["primary"] - cash["financing_fees"],
        "minimum_base_cash": float(frame["ending_cash"].min()),
        "minimum_downside_cash": float(frame["ending_cash_downside"].min()),
        "ipo_target_date": ipo_target,
        "cash_at_ipo": float(ipo_row["ending_cash"]),
        "downside_cash_at_ipo": float(ipo_row["ending_cash_downside"]),
        "next_equity_need": "None before IPO in the base case",
    }
    return frame, summary


def _build_three_statements(
    assumptions: dict[str, Any],
    financials: pd.DataFrame,
    monthly_cash_flow: pd.DataFrame,
    transaction: dict[str, float],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    years = list(range(2025, 2032))
    income_statement = financials.loc[
        years,
        [
            "revenue",
            "revenue_growth",
            "cost_of_revenue",
            "gross_profit",
            "gross_margin",
            "r_and_d",
            "sales_and_marketing",
            "g_and_a",
            "adjusted_operating_income",
            "adjusted_operating_margin",
            "stock_compensation",
        ],
    ].copy()
    income_statement["net_income"] = income_statement[
        "adjusted_operating_income"
    ]

    cash_settings = assumptions["cash"]
    prior_financing_by_year = {
        int(year): float(value)
        for year, value in cash_settings["prior_financing_by_year"].items()
    }
    cash_flow_statement = pd.DataFrame(index=pd.Index(years, name="year"))
    cash_flow_statement["net_income"] = income_statement["net_income"]
    cash_flow_statement["stock_compensation"] = income_statement[
        "stock_compensation"
    ]
    cash_flow_statement["change_in_working_capital"] = -financials.loc[
        years, "working_capital_change"
    ]
    cash_flow_statement["cash_from_operations"] = (
        cash_flow_statement["net_income"]
        + cash_flow_statement["stock_compensation"]
        + cash_flow_statement["change_in_working_capital"]
    )
    cash_flow_statement["capital_expenditures"] = -financials.loc[years, "capex"]
    cash_flow_statement["free_cash_flow"] = (
        cash_flow_statement["cash_from_operations"]
        + cash_flow_statement["capital_expenditures"]
    )
    cash_flow_statement["prior_equity_financing"] = [
        prior_financing_by_year.get(year, 0.0) for year in years
    ]
    cash_flow_statement["series_e_primary"] = [
        transaction["primary"] if year == 2026 else 0.0 for year in years
    ]
    cash_flow_statement["financing_fees"] = [
        -float(cash_settings["financing_fees"]) if year == 2026 else 0.0
        for year in years
    ]
    cash_flow_statement["net_change_in_cash"] = (
        cash_flow_statement["free_cash_flow"]
        + cash_flow_statement["prior_equity_financing"]
        + cash_flow_statement["series_e_primary"]
        + cash_flow_statement["financing_fees"]
    )

    beginning_cash = float(cash_settings["beginning_cash_2025"])
    ending_cash: list[float] = []
    beginning_cash_values: list[float] = []
    for year in years:
        beginning_cash_values.append(beginning_cash)
        beginning_cash += cash_flow_statement.loc[year, "net_change_in_cash"]
        ending_cash.append(beginning_cash)
    cash_flow_statement["beginning_cash"] = beginning_cash_values
    cash_flow_statement["ending_cash"] = ending_cash

    balance_sheet = pd.DataFrame(index=pd.Index(years, name="year"))
    balance_sheet["cash"] = cash_flow_statement["ending_cash"]
    balance_sheet["accounts_receivable"] = financials.loc[years, "revenue"] * 0.10
    balance_sheet["other_current_assets"] = financials.loc[years, "revenue"] * 0.04

    property_and_equipment: list[float] = []
    prior_property_and_equipment = 30.0
    for year in years:
        prior_property_and_equipment += financials.loc[year, "capex"]
        property_and_equipment.append(prior_property_and_equipment)
    balance_sheet["property_and_equipment"] = property_and_equipment
    balance_sheet["other_assets"] = financials.loc[years, "revenue"] * 0.03
    balance_sheet["total_assets"] = balance_sheet[
        [
            "cash",
            "accounts_receivable",
            "other_current_assets",
            "property_and_equipment",
            "other_assets",
        ]
    ].sum(axis=1)

    balance_sheet["accounts_payable"] = financials.loc[years, "cost_of_revenue"] * 0.05
    balance_sheet["deferred_revenue"] = financials.loc[years, "revenue"] * 0.02
    balance_sheet["other_liabilities"] = financials.loc[years, "revenue"] * 0.01
    target_net_working_capital = financials.loc[
        years, "working_capital_change"
    ].cumsum()
    balance_sheet["accrued_expenses"] = (
        balance_sheet["accounts_receivable"]
        + balance_sheet["other_current_assets"]
        - balance_sheet["accounts_payable"]
        - balance_sheet["deferred_revenue"]
        - balance_sheet["other_liabilities"]
        - target_net_working_capital
    ).clip(lower=0)
    balance_sheet["total_debt"] = 0.0
    balance_sheet["total_liabilities"] = balance_sheet[
        [
            "accounts_payable",
            "accrued_expenses",
            "deferred_revenue",
            "other_liabilities",
            "total_debt",
        ]
    ].sum(axis=1)
    balance_sheet["total_equity"] = (
        balance_sheet["total_assets"] - balance_sheet["total_liabilities"]
    )
    balance_sheet["liabilities_and_equity"] = (
        balance_sheet["total_liabilities"] + balance_sheet["total_equity"]
    )
    balance_sheet["balance_check"] = (
        balance_sheet["total_assets"]
        - balance_sheet["liabilities_and_equity"]
    )

    monthly_2026_ending_cash = monthly_cash_flow.loc[
        monthly_cash_flow.index.year == 2026, "ending_cash"
    ].iloc[-1]
    if round(monthly_2026_ending_cash, 6) != round(
        cash_flow_statement.loc[2026, "ending_cash"],
        6,
    ):
        raise ValueError("Monthly and annual 2026 ending cash do not reconcile")

    return income_statement, cash_flow_statement, balance_sheet


def build_model(path: Path = ASSUMPTIONS_PATH) -> ModelBundle:
    assumptions = load_yaml(path)
    financials = _build_financials(assumptions)
    scenarios = _build_scenarios(assumptions, financials)
    transaction = _build_transaction(assumptions)
    monthly_cash_flow, cash_summary = _build_monthly_cash_flow(
        assumptions,
        financials,
        scenarios,
        transaction,
    )
    income_statement, cash_flow_statement, balance_sheet = (
        _build_three_statements(
            assumptions,
            financials,
            monthly_cash_flow,
            transaction,
        )
    )
    return ModelBundle(
        assumptions=assumptions,
        financials=financials,
        scenarios=scenarios,
        market=_build_market(assumptions),
        use_of_proceeds=_build_use_of_proceeds(assumptions),
        unit_economics=_build_unit_economics(assumptions),
        cohorts=_build_cohorts(assumptions),
        transaction=transaction,
        monthly_cash_flow=monthly_cash_flow,
        cash_summary=cash_summary,
        income_statement=income_statement,
        cash_flow_statement=cash_flow_statement,
        balance_sheet=balance_sheet,
    )
