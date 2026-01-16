import os

import mlflow
import mlflow.pyfunc
import pandas as pd
import requests
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


def resolve_registered_model_version(model_name: str, model_info: object) -> int | None:
    version_value = getattr(model_info, "registered_model_version", None)
    if version_value is not None:
        try:
            return int(version_value)
        except (TypeError, ValueError):
            pass
    try:
        client = mlflow.tracking.MlflowClient()
        versions = client.get_latest_versions(model_name)
        if versions:
            return max(int(v.version) for v in versions if v.version is not None)
    except mlflow.exceptions.MlflowException:
        return None
    return None


def resolve_domino_url() -> tuple[str, str]:
    return "https://se-demo.domino.tech:443", "DOMINO_URL"


def resolve_environment_id(
    domino_url: str,
    api_key: str,
    environment_id: str,
    project_id: str,
) -> tuple[str | None, str]:
    if environment_id:
        return environment_id, "DOMINO_ENVIRONMENT_ID"

    if not (domino_url and api_key):
        return None, "missing DOMINO_URL/DOMINO_USER_API_KEY"

    if project_id:
        headers = {"X-Domino-Api-Key": api_key}
        try:
            response = requests.get(
                f"{domino_url}/v4/projects/{project_id}/settings",
                headers=headers,
                timeout=30,
            )
            response.raise_for_status()
        except requests.RequestException:
            response = None
        if response is not None:
            content_type = response.headers.get("content-type", "")
            if "application/json" in content_type:
                try:
                    payload = response.json()
                except requests.JSONDecodeError:
                    payload = None
                if isinstance(payload, dict):
                    default_env = payload.get("defaultEnvironmentId")
                    if default_env:
                        return default_env, "project settings"

    env_name = os.environ.get("DOMINO_ENVIRONMENT_NAME", "")
    headers = {"X-Domino-Api-Key": api_key}
    try:
        response = requests.get(
            f"{domino_url}/api/environments/beta/environments",
            params={"limit": 100},
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
    except requests.RequestException:
        return None

    payload = response.json()
    environments = payload.get("environments", [])
    if env_name:
        for env in environments:
            if env.get("name") == env_name and not env.get("archived"):
                return env.get("id"), "DOMINO_ENVIRONMENT_NAME"
    for env in environments:
        if not env.get("archived"):
            return env.get("id"), "fallback environment list"
    return None, "no environments found"


def register_model_api_endpoint(
    model_api_name: str,
    registered_model_name: str,
    registered_model_version: int,
) -> None:
    domino_url, domino_url_source = resolve_domino_url()
    domino_url = domino_url.rstrip("/")
    api_key = os.environ.get("DOMINO_USER_API_KEY", "")
    project_id = os.environ.get("DOMINO_PROJECT_ID", "")
    environment_id = os.environ.get("DOMINO_ENVIRONMENT_ID", "")
    resolved_environment_id, env_source = resolve_environment_id(
        domino_url,
        api_key,
        environment_id,
        project_id,
    )

    if not (domino_url and api_key and project_id and resolved_environment_id):
        missing = [
            name
            for name, value in {
                "DOMINO_URL": domino_url,
                "DOMINO_USER_API_KEY": api_key,
                "DOMINO_PROJECT_ID": project_id,
                "DOMINO_ENVIRONMENT_ID": resolved_environment_id,
            }.items()
            if not value
        ]
        print(f"Skipping model API registration; missing: {', '.join(missing)}")
        return

    print(
        "Using Domino URL from "
        f"{domino_url_source} and environment from {env_source}."
    )
    payload = {
        "name": model_api_name,
        "description": f"Model API for registered model {registered_model_name} v{registered_model_version}",
        "environmentId": resolved_environment_id,
        "isAsync": False,
        "strictNodeAntiAffinity": False,
        "environmentVariables": [],
        "version": {
            "projectId": project_id,
            "source": {
                "type": "Registry",
                "registeredModelName": registered_model_name,
                "registeredModelVersion": registered_model_version,
            },
            "logHttpRequestResponse": True,
            "monitoringEnabled": True,
        },
    }

    headers = {"X-Domino-Api-Key": api_key}
    response = requests.post(
        f"{domino_url}/api/modelServing/v1/modelApis",
        json=payload,
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()
    model_api = response.json()
    model_api_id = model_api.get("id", "unknown")
    print(f"Registered model API endpoint {model_api_name} (id={model_api_id})")


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

def normalize_endpoint_name(name: str) -> str:
    if not any(ch for ch in name if not ch.isalnum()):
        return name
    parts = []
    current = []
    for ch in name:
        if ch.isalnum():
            current.append(ch)
        elif current:
            parts.append("".join(current))
            current = []
    if current:
        parts.append("".join(current))
    return "".join(part[:1].upper() + part[1:] for part in parts if part)


registered_model_version = resolve_registered_model_version(
    "rihedgingsk-model",
    model_info,
)
if registered_model_version is not None:
    endpoint_name = normalize_endpoint_name("rihedgingsk-model-api")
    register_model_api_endpoint(
        model_api_name=endpoint_name,
        registered_model_name="rihedgingsk-model",
        registered_model_version=registered_model_version,
    )
else:
    print("Skipping model API registration; unable to resolve registered model version.")
