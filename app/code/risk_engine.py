"""
risk_engine.py

Capital-preserving risk engine prototype for PTR.
This module NEVER places trades. It only decides:
APPROVE / REJECT / SKIP, with a reason.
"""

from datetime import datetime, date
from typing import Dict, Any


class RiskDecision:
    APPROVE = "APPROVE"
    REJECT = "REJECT"
    SKIP = "SKIP"


class RiskResult:
    def __init__(self, decision: str, reason: str):
        self.decision = decision
        self.reason = reason

    def __repr__(self) -> str:
        return f"{self.decision}: {self.reason}"


class RiskEngine:
    """
    Hard‑rule risk engine.
    All numbers are intentionally conservative.
    """

    # ===== GLOBAL SAFETY LIMITS =====
    MAX_SINGLE_TRADE_CASH_PCT = 0.20     # never risk >20% of cash
    MAX_TICKER_EQUITY_PCT = 0.10          # max 10% equity per ticker
    MAX_TOTAL_EQUITY_PCT = 0.30           # max 30% invested overall
    MAX_POSITION_GROWTH_PCT = 0.10        # max +10% add to existing position
    MAX_TRADE_AGE_DAYS = 7                # ignore stale trades
    STOP_LOSS_PCT = 0.02                  # mandatory 2% stop
    ENABLE_TRADING = True                 # global kill switch

    def __init__(self, account_state: Dict[str, Any]):
        """
        account_state must include:
        - cash
        - equity
        - positions: dict[symbol] -> {qty, market_value}
        """
        self.cash = float(account_state["cash"])
        self.equity = float(account_state["equity"])
        self.positions = account_state.get("positions", {})

    # ---------- PUBLIC ENTRY POINT ----------

    def evaluate_trade(self, trade: Dict[str, Any]) -> RiskResult:
        """
        trade must include:
        - symbol
        - side: BUY / SELL
        - transaction_date (date)
        - est_price
        """

        if not self.ENABLE_TRADING:
            return RiskResult(RiskDecision.REJECT, "Trading globally disabled")

        symbol = trade["symbol"]
        side = trade["side"].upper()
        est_price = float(trade["est_price"])
        txn_date: date = trade["transaction_date"]

        # 1) Freshness gate
        age_days = (date.today() - txn_date).days
        if age_days > self.MAX_TRADE_AGE_DAYS:
            return RiskResult(RiskDecision.SKIP, "Trade is stale")

        # 2) Estimate order size (small probe only)
        max_trade_value = min(
            self.cash * 0.05,
            self.equity * 0.02
        )

        if max_trade_value <= 0:
            return RiskResult(RiskDecision.REJECT, "No available capital")

        order_qty = max_trade_value / est_price
        order_cost = order_qty * est_price

        # 3) Cash safety
        if order_cost > self.cash * self.MAX_SINGLE_TRADE_CASH_PCT:
            return RiskResult(
                RiskDecision.REJECT,
                "Order exceeds cash risk limit",
            )

        # 4) Existing position rules
        if symbol in self.positions:
            pos_value = float(self.positions[symbol]["market_value"])
            new_value = pos_value + order_cost

            if side == "BUY":
                if new_value > pos_value * (1 + self.MAX_POSITION_GROWTH_PCT):
                    return RiskResult(
                        RiskDecision.REJECT,
                        "Position growth limit exceeded",
                    )

            if side == "SELL":
                return RiskResult(
                    RiskDecision.APPROVE,
                    "Risk‑reducing sell approved",
                )

        # 5) Portfolio exposure limit
        total_invested_after = self._total_invested() + order_cost
        if total_invested_after > self.equity * self.MAX_TOTAL_EQUITY_PCT:
            return RiskResult(
                RiskDecision.REJECT,
                "Total portfolio exposure limit exceeded",
            )

        # 6) Mandatory stop‑loss validation
        stop_price = est_price * (1 - self.STOP_LOSS_PCT)
        if stop_price <= 0:
            return RiskResult(
                RiskDecision.REJECT,
                "Invalid stop‑loss calculation",
            )

        return RiskResult(
            RiskDecision.APPROVE,
            "Trade approved under conservative limits",
        )

    # ---------- HELPERS ----------

    def _total_invested(self) -> float:
        return sum(float(p["market_value"]) for p in self.positions.values())
