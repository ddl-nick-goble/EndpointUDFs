from typing import List, Tuple

import json

import mlflow.pyfunc
import numpy as np
import pandas as pd

from credit_curve_model import RATINGS, _spreads

REQUIRED_COLS = [
    "loan_id",
    "pd_1y",
    "ead",
    "lgd",
    "rating",
    "years_to_maturity",
    "curve_tenors",
    "curve_rates",
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


def _coerce_curve_array(value, label: str) -> np.ndarray:
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError as exc:
            raise ValueError(f"Invalid {label} JSON array") from exc
    if value is None:
        raise ValueError(f"Missing {label}")
    return np.asarray(value, dtype=float)


def _curve_from_arrays(curve_tenors, curve_rates) -> pd.DataFrame:
    tenors = _coerce_curve_array(curve_tenors, "curve_tenors")
    rates = _coerce_curve_array(curve_rates, "curve_rates")
    if tenors.shape != rates.shape:
        raise ValueError("curve_tenors and curve_rates must have matching lengths")
    curve_date = "static"
    data = {"years": tenors, "risk_free_rate": rates}
    for rating in RATINGS:
        data[f"spread_{rating}"] = [
            _spreads(float(years), curve_date)[rating] for years in tenors
        ]
    return pd.DataFrame(data).sort_values("years")


class ExpectedLossModel(mlflow.pyfunc.PythonModel):
    def predict(self, context, model_input: pd.DataFrame) -> pd.DataFrame:
        _ensure_columns(model_input, REQUIRED_COLS)
        _coerce_numeric(model_input, ["pd_1y", "ead", "lgd", "years_to_maturity"])

        curve_tenors = model_input["curve_tenors"].iloc[0]
        curve_rates = model_input["curve_rates"].iloc[0]
        curve_df = _curve_from_arrays(curve_tenors, curve_rates)

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
