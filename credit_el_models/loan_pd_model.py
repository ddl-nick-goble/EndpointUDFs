from typing import List

import mlflow.pyfunc
import numpy as np
import pandas as pd

PURPOSES = ["purchase", "refi", "cash_out", "other"]
REQUIRED_COLS = [
    "loan_id",
    "fico",
    "dti",
    "ltv",
    "loan_age_months",
    "original_balance",
    "interest_rate",
    "employment_length_years",
    "delinquency_30d_12m",
    "loan_purpose",
]


def _ensure_columns(model_input: pd.DataFrame, columns: List[str]) -> None:
    missing = set(columns) - set(model_input.columns)
    if missing:
        missing_list = ", ".join(sorted(missing))
        raise ValueError(f"Missing required columns: {missing_list}")


def _coerce_numeric(df: pd.DataFrame, columns: List[str]) -> None:
    for col in columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")


def _encode_purpose(series: pd.Series) -> np.ndarray:
    categories = pd.Categorical(series, categories=PURPOSES)
    return categories.codes


class LoanPDModel(mlflow.pyfunc.PythonModel):
    def load_context(self, context):
        import xgboost as xgb

        self.model = xgb.XGBClassifier()
        self.model.load_model(context.artifacts["xgb_model"])

    def _extract_features(self, model_input: pd.DataFrame) -> pd.DataFrame:
        _ensure_columns(model_input, REQUIRED_COLS)
        features = model_input.copy()
        _coerce_numeric(
            features,
            [
                "fico",
                "dti",
                "ltv",
                "loan_age_months",
                "original_balance",
                "interest_rate",
                "employment_length_years",
                "delinquency_30d_12m",
            ],
        )
        features["loan_purpose_code"] = _encode_purpose(features["loan_purpose"])
        return features[
            [
                "fico",
                "dti",
                "ltv",
                "loan_age_months",
                "original_balance",
                "interest_rate",
                "employment_length_years",
                "delinquency_30d_12m",
                "loan_purpose_code",
            ]
        ]

    def predict(self, context, model_input: pd.DataFrame) -> pd.DataFrame:
        features = self._extract_features(model_input)
        pds = self.model.predict_proba(features)[:, 1]
        return pd.DataFrame(
            {
                "loan_id": model_input["loan_id"].astype(str).values,
                "pd_1y": pds,
            }
        )
