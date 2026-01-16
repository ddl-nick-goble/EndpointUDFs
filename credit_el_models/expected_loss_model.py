from typing import List, Tuple

import mlflow.pyfunc
import numpy as np
import pandas as pd

from credit_curve_model import json_to_curve

REQUIRED_COLS = [
    "loan_id",
    "pd_1y",
    "ead",
    "lgd",
    "rating",
    "years_to_maturity",
    "curve_json",
]

RISK_WEIGHTS = {
    "AAA": 0.20,
    "AA": 0.20,
    "A": 0.50,
    "BBB": 1.00,
    "BB": 1.50,
    "B": 2.00,
}


def _ensure_columns(model_input: pd.DataFrame, columns: List[str]) -> None:
    missing = set(columns) - set(model_input.columns)
    if missing:
        missing_list = ", ".join(sorted(missing))
        raise ValueError(f"Missing required columns: {missing_list}")


def _coerce_numeric(model_input: pd.DataFrame, columns: List[str]) -> None:
    for col in columns:
        model_input[col] = pd.to_numeric(model_input[col], errors="coerce")


def get_risky_discount_factor(
    curve_df: pd.DataFrame,
    rating: str,
    years: float,
) -> float:
    spread_col = f"spread_{rating}"
    if spread_col not in curve_df.columns:
        raise ValueError(f"Unsupported rating: {rating}")
    rf_rate = float(np.interp(years, curve_df["years"], curve_df["risk_free_rate"]))
    spread = float(np.interp(years, curve_df["years"], curve_df[spread_col]))
    risky_rate = rf_rate + spread
    return float(np.exp(-risky_rate * years))


def compute_expected_loss(row: pd.Series, curve_df: pd.DataFrame) -> Tuple[float, float, float]:
    el_undisc = row["pd_1y"] * row["lgd"] * row["ead"]
    df = get_risky_discount_factor(curve_df, row["rating"], row["years_to_maturity"])
    el_disc = el_undisc * df
    rwa = row["ead"] * RISK_WEIGHTS.get(row["rating"], 1.0)
    return float(el_undisc), float(el_disc), float(rwa)


class ExpectedLossModel(mlflow.pyfunc.PythonModel):
    def predict(self, context, model_input: pd.DataFrame) -> pd.DataFrame:
        _ensure_columns(model_input, REQUIRED_COLS)
        _coerce_numeric(model_input, ["pd_1y", "ead", "lgd", "years_to_maturity"])

        curve_json = model_input["curve_json"].iloc[0]
        curve_df = json_to_curve(curve_json).sort_values("years")

        results = []
        for _, row in model_input.iterrows():
            el_u, el_d, rwa = compute_expected_loss(row, curve_df)
            risk_weight = RISK_WEIGHTS.get(row["rating"], 1.0)
            results.append(
                {
                    "loan_id": row["loan_id"],
                    "ead": row["ead"],
                    "pd_1y": row["pd_1y"],
                    "lgd": row["lgd"],
                    "el_undiscounted": el_u,
                    "el_discounted": el_d,
                    "rwa": rwa,
                    "risk_weight": risk_weight,
                }
            )

        totals = {
            "loan_id": "_TOTAL",
            "ead": float(model_input["ead"].sum()),
            "pd_1y": None,
            "lgd": None,
            "el_undiscounted": float(sum(row["el_undiscounted"] for row in results)),
            "el_discounted": float(sum(row["el_discounted"] for row in results)),
            "rwa": float(sum(row["rwa"] for row in results)),
            "risk_weight": None,
        }
        results.append(totals)

        return pd.DataFrame(results)
