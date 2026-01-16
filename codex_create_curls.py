import html
import json
import os
import re
from pathlib import Path

import requests

DOMINO_URL_ENV = "DOMINO_URL"
DOMINO_USER_HOST_ENV = "DOMINO_USER_HOST"
DOMINO_API_HOST_ENV = "DOMINO_API_HOST"
DOMINO_API_PROXY_ENV = "DOMINO_API_PROXY"
DOMINO_API_KEY_ENV = "DOMINO_USER_API_KEY"
DOMINO_PROJECT_ID_ENV = "DOMINO_PROJECT_ID"
MLFLOW_TRACKING_URI_ENV = "MLFLOW_TRACKING_URI"
DEFAULT_DOMINO_URL = "https://se-demo.domino.tech:443"


def read_env(name: str) -> str:
    return os.environ.get(name, "").strip()


def resolve_domino_url() -> str:
    for name in (DOMINO_URL_ENV, DOMINO_USER_HOST_ENV, DOMINO_API_HOST_ENV, DOMINO_API_PROXY_ENV):
        value = read_env(name)
        if value:
            return value
    return DEFAULT_DOMINO_URL


def get_models(domino_url: str, api_key: str, project_id: str) -> list[dict]:
    url = f"{domino_url.rstrip('/')}/v4/modelManager/getModels"
    headers = {"X-Domino-Api-Key": api_key}
    resp = requests.get(url, params={"projectId": project_id}, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json()


def fetch_overview_curl(domino_url: str, api_key: str, model_id: str) -> str | None:
    url = f"{domino_url.rstrip('/')}/models/{model_id}/overview"
    headers = {"X-Domino-Api-Key": api_key}
    resp = requests.get(url, headers=headers, timeout=15)
    if resp.status_code != 200:
        return None
    match = re.search(
        r'<div role="tabpanel" class="tab-pane" id="language-curl">.*?<pre[^>]*>(.*?)</pre>',
        resp.text,
        flags=re.S | re.I,
    )
    if not match:
        return None
    return html.unescape(match.group(1)).strip()


def get_mlflow_model_source(tracking_uri: str, model_name: str, model_version: int) -> str | None:
    url = f"{tracking_uri.rstrip('/')}/api/2.0/mlflow/model-versions/get"
    resp = requests.get(url, params={"name": model_name, "version": model_version}, timeout=15)
    resp.raise_for_status()
    source = resp.json().get("model_version", {}).get("source")
    return source if isinstance(source, str) and source else None


def load_mlflow_input_example(source: str) -> dict | None:
    try:
        from mlflow import artifacts
    except Exception:
        return None
    os.environ.setdefault("MLFLOW_ENABLE_ARTIFACTS_PROGRESS_BAR", "false")
    local_dir = artifacts.download_artifacts(artifact_uri=source)
    for filename in ("serving_input_example.json", "input_example.json"):
        path = Path(local_dir) / filename
        if path.exists():
            return json.loads(path.read_text())
    return None


def example_to_payload(example: object) -> dict | None:
    if not isinstance(example, dict):
        return None
    if "data" in example:
        return example
    dataframe_split = example.get("dataframe_split")
    if not isinstance(dataframe_split, dict):
        return None
    columns = dataframe_split.get("columns")
    data = dataframe_split.get("data")
    if not isinstance(columns, list) or not isinstance(data, list) or not data:
        return None
    row = data[0]
    if not isinstance(row, (list, tuple)) or len(row) != len(columns):
        return None
    return {"data": {str(col): row[idx] for idx, col in enumerate(columns)}}


def replace_curl_data(curl_text: str, payload: dict) -> str:
    payload_json = json.dumps(payload)
    replacement = f"-d '{payload_json}'"
    lines = curl_text.splitlines()
    for idx, line in enumerate(lines):
        if line.lstrip().startswith("-d "):
            indent = re.match(r"\s*", line).group(0)
            suffix_match = re.search(r"\s*\\\s*$", line)
            suffix = suffix_match.group(0) if suffix_match else ""
            lines[idx] = f"{indent}{replacement}{suffix}"
            return "\n".join(lines)
    insert_at = len(lines)
    for idx, line in enumerate(lines):
        if line.lstrip().startswith("-u "):
            insert_at = idx
            break
    lines.insert(insert_at, replacement)
    return "\n".join(lines)


def main() -> None:
    domino_url = resolve_domino_url()
    api_key = read_env(DOMINO_API_KEY_ENV)
    project_id = read_env(DOMINO_PROJECT_ID_ENV)
    tracking_uri = read_env(MLFLOW_TRACKING_URI_ENV)
    if not api_key or not project_id:
        raise SystemExit("Missing DOMINO_USER_API_KEY or DOMINO_PROJECT_ID.")

    models = get_models(domino_url, api_key, project_id)
    cache: dict[tuple[str, int], dict | None] = {}

    print("== Model inference curls ==")
    for model in models:
        model_id = model.get("id")
        name = model.get("name") or model_id
        active = model.get("activeVersion") or {}
        registered_name = active.get("registeredModelName")
        registered_version = active.get("registeredModelVersion")
        if not model_id:
            continue

        curl = fetch_overview_curl(domino_url, api_key, model_id)
        if not curl:
            continue

        payload = None
        if tracking_uri and registered_name and isinstance(registered_version, int):
            key = (registered_name, registered_version)
            if key not in cache:
                try:
                    source = get_mlflow_model_source(
                        tracking_uri, registered_name, registered_version
                    )
                    example = load_mlflow_input_example(source) if source else None
                    cache[key] = example_to_payload(example)
                except requests.RequestException:
                    cache[key] = None
            payload = cache.get(key)

        if payload:
            curl = replace_curl_data(curl, payload)

        print(f"- {name}")
        print(curl)
        print()


if __name__ == "__main__":
    main()
