from pathlib import Path

import pandas as pd

from credit_curve_model import CreditCurveModel, curve_to_json
from expected_loss_model import ExpectedLossModel
from loan_pd_model import LoanPDModel
from train_pd_model import train_and_save_pd_model


class _DummyContext:
    def __init__(self, artifacts):
        self.artifacts = artifacts


def run_smoke_test():
    curve_in = pd.DataFrame({"curve_date": ["2024-12-31"]})
    curve_out = CreditCurveModel().predict(None, curve_in)
    curve_json = curve_to_json(curve_out)

    model_path = Path("pd_model.json")
    if not model_path.exists():
        train_and_save_pd_model(str(model_path))

    loans_in = pd.DataFrame(
        {
            "loan_id": ["L001", "L002"],
            "fico": [720, 650],
            "dti": [0.35, 0.48],
            "ltv": [0.8, 0.92],
            "loan_age_months": [12, 36],
            "original_balance": [250000, 180000],
            "interest_rate": [0.065, 0.075],
            "employment_length_years": [5, 2],
            "delinquency_30d_12m": [0, 1],
            "loan_purpose": ["purchase", "refi"],
        }
    )
    pd_model = LoanPDModel()
    pd_model.load_context(_DummyContext({"xgb_model": str(model_path)}))
    pd_out = pd_model.predict(None, loans_in)

    el_in = loans_in.copy()
    el_in["pd_1y"] = pd_out["pd_1y"]
    el_in["ead"] = [235000, 168000]
    el_in["lgd"] = [0.40, 0.45]
    el_in["rating"] = ["A", "BBB"]
    el_in["years_to_maturity"] = [4.5, 2.0]
    el_in["curve_json"] = curve_json

    el_out = ExpectedLossModel().predict(None, el_in)
    print(el_out)


if __name__ == "__main__":
    run_smoke_test()
