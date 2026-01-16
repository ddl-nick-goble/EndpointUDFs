import mlflow
import mlflow.pyfunc
import numpy as np
import pandas as pd
from datetime import datetime
from mlflow.models import infer_signature

TENORS = [
    ("3M", 0.25),
    ("1Y", 1.0),
    ("2Y", 2.0),
    ("3Y", 3.0),
    ("5Y", 5.0),
    ("7Y", 7.0),
    ("10Y", 10.0),
    ("15Y", 15.0),
    ("20Y", 20.0),
    ("30Y", 30.0),
]


def get_asset_yield_curve(curve_id: str, curve_date: str) -> pd.DataFrame:
    """
    curve_id: 'treasury_zero' | 'SOFR'
    curve_date: YYYY-MM-DD
    """
    curve_date = pd.to_datetime(curve_date)

    if curve_id == "treasury_zero":
        base_rate = 0.03
        slope = 0.002
    elif curve_id == "SOFR":
        base_rate = 0.035
        slope = 0.0015
    else:
        raise ValueError(f"Unsupported curve_id: {curve_id}")

    rows = []
    for tenor, years in TENORS:
        rate = base_rate + slope * years
        rows.append({
            "tenor": tenor,
            "date": curve_date,
            "rate": round(rate, 4),
        })

    return pd.DataFrame(rows)


def get_policy_cashflows(policy_ids: list[str]) -> pd.DataFrame:
    """
    Returns a DataFrame:
      index: dates
      columns: policy_ids
      values: expected cashflows
    """
    start_date = pd.Timestamp("2025-01-01")
    periods = 10  # 10 years annual cashflows
    dates = pd.date_range(start=start_date, periods=periods, freq="YE")

    data = {}
    for pid in policy_ids:
        notional = int(1_000_000 + (hash(pid) % 250_000))
        annual_cf = notional // periods
        data[pid] = [annual_cf] * periods

    df = pd.DataFrame(data, index=dates, dtype="int64")
    df.index.name = "date"
    return df


def discount_liabilities(
    cashflows: pd.DataFrame,
    yield_curve: pd.DataFrame,
) -> pd.DataFrame:
    """
    cashflows:
      index: dates
      columns: policy_ids

    yield_curve:
      columns: tenor, date, rate
    """
    tenor_years = {
        "3M": 0.25, "1Y": 1, "2Y": 2, "3Y": 3, "5Y": 5,
        "7Y": 7, "10Y": 10, "15Y": 15, "20Y": 20, "30Y": 30,
    }

    curve = yield_curve.copy()
    curve["years"] = curve["tenor"].map(tenor_years)
    curve = curve.sort_values("years")

    base_date = curve["date"].iloc[0]

    discounted = cashflows.astype("float64")

    for date in discounted.index:
        t = (date - base_date).days / 365.25
        nearest = curve.iloc[(curve["years"] - t).abs().argsort().iloc[0]]
        rate = nearest["rate"]
        df = (1 + rate) ** t
        discounted.loc[date] = discounted.loc[date] / df

    return discounted


def price_with_capital_cost(
    discounted_liabilities: pd.DataFrame,
    capital_ratio: float,
    margin: float,
) -> pd.DataFrame:
    """
    discounted_liabilities:
      index: dates
      columns: policy_ids
    """
    pv = discounted_liabilities.sum(axis=0)

    price = pv * (1 + capital_ratio)
    price = price * (1 + margin)

    return pd.DataFrame({
        "policy_id": price.index,
        "price": price.values,
    })


class PricingModel(mlflow.pyfunc.PythonModel):
    def predict(self, context, model_input: pd.DataFrame) -> pd.DataFrame:
        discount_model = DiscountModel()
        discounted = discount_model.predict(context, model_input)
        return price_with_capital_cost(
            discounted_liabilities=discounted,
            capital_ratio=0.08,
            margin=0.05,
        )


def _first_value(series: pd.Series):
    value = series.iloc[0]
    if isinstance(value, (list, tuple, np.ndarray, pd.Series)):
        return value[0]
    return value


def _column_as_list(model_input: pd.DataFrame, column: str) -> list:
    if column not in model_input.columns:
        raise ValueError(f"Missing required column: {column}")
    series = model_input[column]
    if len(series) == 1 and isinstance(series.iloc[0], (list, tuple, np.ndarray, pd.Series)):
        return list(series.iloc[0])
    return series.tolist()


class YieldCurveModel(mlflow.pyfunc.PythonModel):
    def predict(self, context, model_input: pd.DataFrame) -> pd.DataFrame:
        curve_id = _first_value(model_input["curve_id"])
        curve_date = _first_value(model_input["curve_date"])
        return get_asset_yield_curve(curve_id=curve_id, curve_date=curve_date)


class CashflowsModel(mlflow.pyfunc.PythonModel):
    def predict(self, context, model_input: pd.DataFrame) -> pd.DataFrame:
        policy_ids = _column_as_list(model_input, "policy_id")
        return get_policy_cashflows(policy_ids=policy_ids)


class DiscountModel(mlflow.pyfunc.PythonModel):
    def predict(self, context, model_input: pd.DataFrame) -> pd.DataFrame:
        curve_id = _first_value(model_input["curve_id"])
        curve_date = _first_value(model_input["curve_date"])
        policy_ids = _column_as_list(model_input, "policy_id")

        curve_input = pd.DataFrame({
            "curve_id": [curve_id],
            "curve_date": [curve_date],
        })
        cashflows_input = pd.DataFrame({"policy_id": policy_ids})

        curve_model = YieldCurveModel()
        cashflows_model = CashflowsModel()
        curve = curve_model.predict(context, curve_input)
        cashflows = cashflows_model.predict(context, cashflows_input)

        return discount_liabilities(
            cashflows=cashflows,
            yield_curve=curve,
        )


yield_curve_input_example = pd.DataFrame({
    "curve_id": ["treasury_zero"],
    "curve_date": ["2024-12-31"],
})
yield_curve_output_example = get_asset_yield_curve("treasury_zero", "2024-12-31")

cashflows_input_example = pd.DataFrame({"policy_id": ["POL001", "POL002"]})
cashflows_output_example = get_policy_cashflows(["POL001", "POL002"]).astype("float64")

discount_input_example = pd.DataFrame({
    "curve_id": ["treasury_zero", "treasury_zero"],
    "curve_date": ["2024-12-31", "2024-12-31"],
    "policy_id": ["POL001", "POL002"],
})
discount_output_example = discount_liabilities(
    cashflows=cashflows_output_example,
    yield_curve=yield_curve_output_example,
)

pricing_input_example = discount_input_example
pricing_output_example = pd.DataFrame({"policy_id": ["POL001"], "price": [120000.0]})

signature_yield_curve = infer_signature(
    yield_curve_input_example,
    yield_curve_output_example,
)
signature_cashflows = infer_signature(
    cashflows_input_example,
    cashflows_output_example,
)
signature_discount = infer_signature(
    discount_input_example,
    discount_output_example,
)
signature_pricing = infer_signature(
    pricing_input_example,
    pricing_output_example,
)

with mlflow.start_run() as run:
    mlflow.pyfunc.log_model(
        name="alm-yield-curve-model",
        python_model=YieldCurveModel(),
        signature=signature_yield_curve,
        input_example=yield_curve_input_example,
        registered_model_name="alm-yield-curve-model",
    )

with mlflow.start_run() as run:
    mlflow.pyfunc.log_model(
        name="alm-cashflows-model",
        python_model=CashflowsModel(),
        signature=signature_cashflows,
        input_example=cashflows_input_example,
        registered_model_name="alm-cashflows-model",
    )

with mlflow.start_run() as run:
    mlflow.pyfunc.log_model(
        name="alm-discount-model",
        python_model=DiscountModel(),
        signature=signature_discount,
        input_example=discount_input_example,
        registered_model_name="alm-discount-model",
    )

with mlflow.start_run() as run:
    mlflow.pyfunc.log_model(
        name="alm-pricing-model",
        python_model=PricingModel(),
        signature=signature_pricing,
        input_example=pricing_input_example,
        registered_model_name="alm-pricing-model",
    )
