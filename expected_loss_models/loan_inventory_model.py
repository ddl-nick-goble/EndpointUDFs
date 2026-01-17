from __future__ import annotations

from typing import List

import mlflow.pyfunc
import pandas as pd


def _lgd_from_rating(credit_rating: str, original_loan_term_years: float) -> float:
    mapping = {
        "AAA": 0.30,
        "AA": 0.32,
        "A": 0.36,
        "BBB": 0.45,
        "BB": 0.52,
        "B": 0.62,
    }
    base = mapping.get(credit_rating, 0.45)
    if original_loan_term_years <= 5:
        adj = -0.02
    elif original_loan_term_years > 15:
        adj = 0.03
    else:
        adj = 0.0
    return max(0.1, min(0.75, base + adj))


def build_loan_inventory(inventory_date: str) -> pd.DataFrame:
    _ = inventory_date
    rows: List[dict] = [
        {
            "loan_id": "L001",
            "credit_score": 780,
            "debt_to_income_ratio": 0.22,
            "loan_to_value_ratio": 0.68,
            "loan_age_months": 10,
            "original_principal_balance": 420000,
            "interest_rate": 0.045,
            "employment_years": 12,
            "delinquency_30d_past_12m": 0,
            "loan_purpose": "purchase",
            "original_loan_term_years": 30,
            "credit_rating": "AAA",
            "remaining_term_years": 29,
            "loss_given_default": _lgd_from_rating("AAA", 30),
        },
        {
            "loan_id": "L002",
            "credit_score": 750,
            "debt_to_income_ratio": 0.28,
            "loan_to_value_ratio": 0.72,
            "loan_age_months": 24,
            "original_principal_balance": 350000,
            "interest_rate": 0.047,
            "employment_years": 9,
            "delinquency_30d_past_12m": 0,
            "loan_purpose": "refi",
            "original_loan_term_years": 30,
            "credit_rating": "AA",
            "remaining_term_years": 28,
            "loss_given_default": _lgd_from_rating("AA", 30),
        },
        {
            "loan_id": "L003",
            "credit_score": 720,
            "debt_to_income_ratio": 0.33,
            "loan_to_value_ratio": 0.78,
            "loan_age_months": 36,
            "original_principal_balance": 300000,
            "interest_rate": 0.052,
            "employment_years": 7,
            "delinquency_30d_past_12m": 0,
            "loan_purpose": "purchase",
            "original_loan_term_years": 25,
            "credit_rating": "A",
            "remaining_term_years": 22,
            "loss_given_default": _lgd_from_rating("A", 25),
        },
        {
            "loan_id": "L004",
            "credit_score": 700,
            "debt_to_income_ratio": 0.38,
            "loan_to_value_ratio": 0.83,
            "loan_age_months": 48,
            "original_principal_balance": 280000,
            "interest_rate": 0.058,
            "employment_years": 6,
            "delinquency_30d_past_12m": 1,
            "loan_purpose": "refi",
            "original_loan_term_years": 20,
            "credit_rating": "BBB",
            "remaining_term_years": 16,
            "loss_given_default": _lgd_from_rating("BBB", 20),
        },
        {
            "loan_id": "L005",
            "credit_score": 690,
            "debt_to_income_ratio": 0.41,
            "loan_to_value_ratio": 0.86,
            "loan_age_months": 60,
            "original_principal_balance": 260000,
            "interest_rate": 0.062,
            "employment_years": 5,
            "delinquency_30d_past_12m": 1,
            "loan_purpose": "purchase",
            "original_loan_term_years": 20,
            "credit_rating": "BBB",
            "remaining_term_years": 15,
            "loss_given_default": _lgd_from_rating("BBB", 20),
        },
        {
            "loan_id": "L006",
            "credit_score": 670,
            "debt_to_income_ratio": 0.45,
            "loan_to_value_ratio": 0.90,
            "loan_age_months": 72,
            "original_principal_balance": 240000,
            "interest_rate": 0.068,
            "employment_years": 4,
            "delinquency_30d_past_12m": 2,
            "loan_purpose": "refi",
            "original_loan_term_years": 15,
            "credit_rating": "BB",
            "remaining_term_years": 9,
            "loss_given_default": _lgd_from_rating("BB", 15),
        },
        {
            "loan_id": "L007",
            "credit_score": 655,
            "debt_to_income_ratio": 0.48,
            "loan_to_value_ratio": 0.92,
            "loan_age_months": 84,
            "original_principal_balance": 210000,
            "interest_rate": 0.072,
            "employment_years": 3,
            "delinquency_30d_past_12m": 2,
            "loan_purpose": "purchase",
            "original_loan_term_years": 15,
            "credit_rating": "B",
            "remaining_term_years": 8,
            "loss_given_default": _lgd_from_rating("B", 15),
        },
        {
            "loan_id": "L008",
            "credit_score": 740,
            "debt_to_income_ratio": 0.30,
            "loan_to_value_ratio": 0.75,
            "loan_age_months": 18,
            "original_principal_balance": 380000,
            "interest_rate": 0.049,
            "employment_years": 8,
            "delinquency_30d_past_12m": 0,
            "loan_purpose": "purchase",
            "original_loan_term_years": 30,
            "credit_rating": "A",
            "remaining_term_years": 28,
            "loss_given_default": _lgd_from_rating("A", 30),
        },
        {
            "loan_id": "L009",
            "credit_score": 705,
            "debt_to_income_ratio": 0.36,
            "loan_to_value_ratio": 0.82,
            "loan_age_months": 30,
            "original_principal_balance": 310000,
            "interest_rate": 0.055,
            "employment_years": 6,
            "delinquency_30d_past_12m": 0,
            "loan_purpose": "refi",
            "original_loan_term_years": 25,
            "credit_rating": "BBB",
            "remaining_term_years": 22,
            "loss_given_default": _lgd_from_rating("BBB", 25),
        },
        {
            "loan_id": "L010",
            "credit_score": 680,
            "debt_to_income_ratio": 0.43,
            "loan_to_value_ratio": 0.88,
            "loan_age_months": 54,
            "original_principal_balance": 230000,
            "interest_rate": 0.064,
            "employment_years": 4,
            "delinquency_30d_past_12m": 1,
            "loan_purpose": "purchase",
            "original_loan_term_years": 20,
            "credit_rating": "BB",
            "remaining_term_years": 14,
            "loss_given_default": _lgd_from_rating("BB", 20),
        },
    ]
    inventory = pd.DataFrame(rows)
    ordered_cols = [
        "loan_id",
        "credit_score",
        "debt_to_income_ratio",
        "loan_to_value_ratio",
        "loan_age_months",
        "original_principal_balance",
        "interest_rate",
        "employment_years",
        "delinquency_30d_past_12m",
        "loan_purpose",
        "credit_rating",
        "original_loan_term_years",
        "remaining_term_years",
        "loss_given_default",
    ]
    return inventory[ordered_cols]


class LoanInventoryModel(mlflow.pyfunc.PythonModel):
    def predict(self, context, model_input: pd.DataFrame) -> pd.DataFrame:
        if "inventory_date" not in model_input.columns:
            raise ValueError("Missing required column: inventory_date")
        inventory_date = str(model_input["inventory_date"].iloc[0])
        return build_loan_inventory(inventory_date)
