from pathlib import Path

import pandas as pd

from credit_curve_model import CreditCurveModel
from expected_loss_model import ExpectedLossModel
from loan_pd_model import LoanPDModel
from train_pd_model import train_and_save_pd_model


class _DummyContext:
    def __init__(self, artifacts):
        self.artifacts = artifacts


def run_smoke_test():
    curve_in = pd.DataFrame({"curve_date": ["2024-12-31"]})
    curve_out = CreditCurveModel().predict(None, curve_in)
    curve_tenors = curve_out["years"].astype(float).tolist()
    curve_rates = curve_out["risk_free_rate"].astype(float).tolist()

    model_path = Path("pd_model.json")
    if not model_path.exists():
        train_and_save_pd_model(str(model_path))

    loans_in = pd.DataFrame(
        {
            "loan_id": ["L001", "L002"],
            "credit_score": [720, 650],
            "debt_to_income_ratio": [0.35, 0.48],
            "loan_to_value_ratio": [0.8, 0.92],
            "loan_age_months": [12, 36],
            "original_principal_balance": [250000, 180000],
            "interest_rate": [0.065, 0.075],
            "employment_years": [5, 2],
            "delinquency_30d_past_12m": [0, 1],
            "loan_purpose": ["purchase", "refi"],
        }
    )
    pd_model = LoanPDModel()
    pd_model.load_context(_DummyContext({"xgb_model": str(model_path)}))
    pd_out = pd_model.predict(None, loans_in)

    el_in = loans_in.copy()
    el_in["probability_of_default_1y"] = pd_out["probability_of_default_1y"]
    el_in["exposure_at_default"] = [235000, 168000]
    el_in["loss_given_default"] = [0.40, 0.45]
    el_in["credit_rating"] = ["A", "BBB"]
    el_in["remaining_term_years"] = [4.5, 2.0]
    el_in["curve_tenors"] = [curve_tenors] * len(el_in)
    el_in["curve_rates"] = [curve_rates] * len(el_in)

    el_out = ExpectedLossModel().predict(None, el_in)
    print(el_out)


if __name__ == "__main__":
    run_smoke_test()
