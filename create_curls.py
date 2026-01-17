import html
import json
import os
import re
import requests

DOMINO_URL = os.environ.get("DOMINO_URL", "https://se-demo.domino.tech:443")
API_KEY = os.environ.get("DOMINO_USER_API_KEY", "")
PROJECT_ID = os.environ.get("DOMINO_PROJECT_ID", "")


def get_models(project_id: str) -> list:
    """Get all models for a project."""
    url = f"{DOMINO_URL}/v4/modelManager/getModels"
    headers = {"X-Domino-Api-Key": API_KEY}
    resp = requests.get(url, params={"projectId": project_id}, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json()


def get_model_signature(model_name: str, model_version: int) -> dict | None:
    """Get the signature for a registered model version from MLflow."""
    tracking_uri = os.environ.get("MLFLOW_TRACKING_URI", "")
    if not tracking_uri:
        return None

    # Get the model version info to find the artifact source
    url = f"{tracking_uri.rstrip('/')}/api/2.0/mlflow/model-versions/get"
    resp = requests.get(url, params={"name": model_name, "version": model_version}, timeout=15)
    if resp.status_code != 200:
        return None

    source = resp.json().get("model_version", {}).get("source")
    if not source:
        return None

    # Download and parse the input example
    try:
        from mlflow import artifacts
        os.environ.setdefault("MLFLOW_ENABLE_ARTIFACTS_PROGRESS_BAR", "false")
        local_dir = artifacts.download_artifacts(artifact_uri=source)

        for fname in ("serving_input_example.json", "input_example.json"):
            path = os.path.join(local_dir, fname)
            if os.path.exists(path):
                with open(path) as f:
                    example = json.load(f)
                # Convert dataframe_split format to simple dict
                if "dataframe_split" in example:
                    cols = example["dataframe_split"]["columns"]
                    row = example["dataframe_split"]["data"][0]
                    return {"data": {col: row[i] for i, col in enumerate(cols)}}
                if "data" in example:
                    return example
    except Exception:
        pass

    return None


def get_curl_from_html(model_id: str) -> str | None:
    """Fetch the model overview page and extract the curl command."""
    url = f"{DOMINO_URL}/models/{model_id}/overview"
    headers = {"X-Domino-Api-Key": API_KEY}
    resp = requests.get(url, headers=headers, timeout=15)
    if resp.status_code != 200:
        return None

    # Extract curl from the HTML
    match = re.search(
        r'<div role="tabpanel" class="tab-pane" id="language-curl">.*?<pre[^>]*>(.*?)</pre>',
        resp.text,
        flags=re.S | re.I,
    )
    if not match:
        return None

    return html.unescape(match.group(1)).strip()


def replace_curl_data(curl_text: str, payload: dict) -> str:
    """Replace the -d data in a curl command with the given payload."""
    payload_json = json.dumps(payload)
    # Match -d followed by single-quoted string containing JSON (may have nested double quotes)
    # Use non-greedy match from opening quote to the closing quote before whitespace/newline or end
    pattern = r"-d\s+'[^']*'"
    replacement = f"-d '{payload_json}'"
    return re.sub(pattern, replacement, curl_text)


def main():
    if not API_KEY or not PROJECT_ID:
        print("Set DOMINO_USER_API_KEY and DOMINO_PROJECT_ID environment variables")
        return

    models = get_models(PROJECT_ID)

    for model in models:
        model_id = model.get("id")
        name = model.get("name")
        active = model.get("activeVersion") or {}
        registered_name = active.get("registeredModelName")
        registered_version = active.get("registeredModelVersion")

        if not model_id:
            continue

        print(f"=== {name} ===")

        # Get the curl from the overview page
        curl = get_curl_from_html(model_id)
        if not curl:
            print(f"  (no curl found in overview page)")
            continue

        # Try to get the signature
        signature = None
        if registered_name and registered_version:
            signature = get_model_signature(registered_name, registered_version)

        if signature:
            # Replace the -d data with the actual signature
            curl = replace_curl_data(curl, signature)
            print(curl)
        else:
            print(curl)
            print(f"  (no signature found for {registered_name}:{registered_version})")

        print()


if __name__ == "__main__":
    main()
