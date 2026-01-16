import mlflow
import mlflow.pyfunc
import pandas as pd
from mlflow.models import infer_signature

mlflow.set_experiment("risk-model-experiment")

X = pd.DataFrame({
    "age": [35.05, 42.3],
    "income": [95000.1, 90000.4],
})

y = pd.DataFrame({
    "risk_score": [0.12, 0.27]
})

class RiskModel(mlflow.pyfunc.PythonModel):
    def predict(self, context, model_input: pd.DataFrame) -> pd.Series:
        # model_input is a pandas DataFrame
        return model_input["age"] * 0.001

signature = infer_signature(X, y)

with mlflow.start_run() as run:
    model_info = mlflow.pyfunc.log_model(
        name="hedging-model",
        python_model=RiskModel(),
        signature=signature,
        input_example=X,
        registered_model_name="rihedgingsk-model"
    )

print(f"Logged model to run {run.info.run_id}: {model_info.model_uri}")
