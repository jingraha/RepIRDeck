from __future__ import annotations

import math

from src.model import build_model


def test_transaction_reconciles() -> None:
    model = build_model()
    tx = model.transaction
    assert tx["primary"] + tx["secondary"] == tx["total_raise"]
    assert tx["post_money"] == tx["pre_money"] + tx["primary"]
    assert math.isclose(tx["primary_ownership"], 800 / 20800)
    assert math.isclose(tx["total_buyer_ownership"], 1000 / 20800)


def test_revenue_segments_reconcile() -> None:
    model = build_model()
    segment_total = model.financials[
        [
            "revenue_individual",
            "revenue_enterprise",
            "revenue_usage_and_deploy",
        ]
    ].sum(axis=1)
    assert segment_total.equals(model.financials["revenue"])


def test_primary_capital_is_fully_allocated() -> None:
    model = build_model()
    assert model.use_of_proceeds["amount"].sum() == model.transaction["primary"]
    assert math.isclose(model.use_of_proceeds["percent"].sum(), 1.0)


def test_scenarios_have_expected_order() -> None:
    model = build_model()
    for year in range(2027, 2032):
        assert (
            model.scenarios["downside"].loc[year, "revenue"]
            < model.scenarios["base"].loc[year, "revenue"]
            < model.scenarios["upside"].loc[year, "revenue"]
        )


def test_market_sums_to_250_billion() -> None:
    model = build_model()
    assert model.market["tam"].sum() == 250_000
    assert model.market["sam"].sum() == 90_000


def test_monthly_cash_model_shows_funding_need_and_runway() -> None:
    model = build_model()
    assert model.cash_summary["funding_need_date"].strftime("%Y-%m") == "2026-11"
    assert (
        model.monthly_cash_flow["primary_financing"].sum()
        == model.transaction["primary"]
    )
    assert model.cash_summary["minimum_base_cash"] >= model.cash_summary["minimum_cash"]
    assert (
        model.cash_summary["minimum_downside_cash"]
        >= model.cash_summary["minimum_cash"]
    )
    assert model.cash_summary["next_equity_need"] == "None before IPO in the base case"


def test_three_statements_reconcile() -> None:
    model = build_model()
    assert model.balance_sheet["balance_check"].abs().max() < 0.000001
    assert (
        model.cash_flow_statement.loc[2026, "ending_cash"]
        == model.monthly_cash_flow.loc["2026-12-31", "ending_cash"]
    )
    assert (
        model.cash_flow_statement["ending_cash"]
        == model.balance_sheet["cash"]
    ).all()
