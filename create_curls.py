import ast
import html
import json
import os
import re
from urllib.parse import urlparse, urlunparse
from pathlib import Path
from urllib.parse import urlencode

import requests

DOMINO_URL_ENV = "DOMINO_URL"
DOMINO_API_HOST_ENV = "DOMINO_API_HOST"
DOMINO_USER_HOST_ENV = "DOMINO_USER_HOST"
DOMINO_API_PROXY_ENV = "DOMINO_API_PROXY"
DOMINO_INFERENCE_HOST_ENV = "DOMINO_INFERENCE_HOST"
DOMINO_PUBLIC_HOST_ENV = "DOMINO_PUBLIC_HOST"
DOMINO_INFERENCE_BASE_URL_ENV = "DOMINO_INFERENCE_BASE_URL"
DOMINO_API_KEY_ENV = "DOMINO_USER_API_KEY"
DOMINO_PROJECT_ID_ENV = "DOMINO_PROJECT_ID"
DOMINO_SIGNATURE_ALIASES_ENV = "DOMINO_SIGNATURE_ALIASES"
MLFLOW_TRACKING_URI_ENV = "MLFLOW_TRACKING_URI"
DEFAULT_DOMINO_URL = "https://se-demo.domino.tech:443"


def read_env(name: str) -> str:
    value = os.environ.get(name, "")
    return value.strip()


def require_env(domino_url: str, api_key: str, project_id: str) -> None:
    missing = []
    if not domino_url:
        missing.append(
            f"{DOMINO_URL_ENV} (or {DOMINO_USER_HOST_ENV}/{DOMINO_API_HOST_ENV}/{DOMINO_API_PROXY_ENV})"
        )
    if not api_key:
        missing.append(DOMINO_API_KEY_ENV)
    if not project_id:
        missing.append(DOMINO_PROJECT_ID_ENV)
    if missing:
        joined = ", ".join(missing)
        raise SystemExit(f"Missing required env vars: {joined}")


def get_json(url: str, api_key: str, params: dict | None = None) -> object:
    headers = {"X-Domino-Api-Key": api_key}
    resp = requests.get(url, params=params, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json()


def build_url(base: str, path: str, params: dict | None = None) -> str:
    url = f"{base.rstrip('/')}{path}"
    if params:
        url = f"{url}?{urlencode(params)}"
    return url


def build_curl(method: str, url: str, api_key: str, data: object | None = None) -> str:
    parts = [
        "curl",
        "-sS",
        "-X",
        method.upper(),
        f'-H "X-Domino-Api-Key: {api_key}"',
        '-H "Accept: application/json"',
    ]
    if data is not None:
        parts.append('-H "Content-Type: application/json"')
        parts.append(f"-d '{json.dumps(data)}'")
    parts.append(f"\"{url}\"")
    return " ".join(parts)


def build_inference_curl(
    url: str, api_key: str, data: object | None = None, url_quote: str = "\""
) -> str:
    parts = [
        "curl",
        f"{url_quote}{url}{url_quote}",
    ]
    if data is not None:
        parts.append("-H 'Content-Type: application/json'")
        parts.append(f"-d '{json.dumps(data)}'")
    parts.append(f"-u {api_key}:{api_key}")
    return " ".join(parts)


def extract_local_signature() -> dict:
    models_path = Path("models.py")
    if not models_path.exists():
        return {}

    try:
        tree = ast.parse(models_path.read_text())
    except SyntaxError:
        return {}

    registered_name = None
    input_dict = None

    for node in ast.walk(tree):
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
            if node.func.attr == "log_model":
                for kw in node.keywords:
                    if kw.arg == "registered_model_name":
                        try:
                            registered_name = ast.literal_eval(kw.value)
                        except ValueError:
                            registered_name = None

    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign):
            continue
        if not any(isinstance(t, ast.Name) and t.id == "X" for t in node.targets):
            continue
        call = node.value
        if not isinstance(call, ast.Call) or not isinstance(call.func, ast.Attribute):
            continue
        if call.func.attr != "DataFrame":
            continue
        dict_node = None
        if call.args:
            dict_node = call.args[0]
        else:
            for kw in call.keywords:
                if kw.arg in (None, "data"):
                    dict_node = kw.value
                    break
        if not isinstance(dict_node, ast.Dict):
            continue
        try:
            input_dict = ast.literal_eval(dict_node)
        except ValueError:
            input_dict = None

    if not registered_name or not input_dict:
        return {}

    columns = list(input_dict.keys())
    rows = list(zip(*[input_dict[c] for c in columns]))
    payload = {"dataframe_split": {"columns": columns, "data": rows}}
    return {registered_name: {"payload": payload, "source": str(models_path)}}


def signature_to_model_payload(signature_payload: dict | None) -> dict | None:
    if not signature_payload:
        return None
    dataframe_split = signature_payload.get("dataframe_split")
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


def example_to_model_payload(example: object) -> dict | None:
    if isinstance(example, dict):
        if "data" in example:
            return example
        return signature_to_model_payload(example)
    return None


def normalize_model_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", name.lower())


def parse_signature_aliases() -> dict[str, str]:
    raw = read_env(DOMINO_SIGNATURE_ALIASES_ENV)
    if not raw:
        return {}
    aliases = {}
    for pair in raw.split(","):
        if "=" not in pair:
            continue
        left, right = pair.split("=", 1)
        left = left.strip()
        right = right.strip()
        if left and right:
            aliases[left] = right
    return aliases


def resolve_signature_payload(
    registered_name: str | None,
    registered_version: int | None,
    signature_map: dict,
    signature_aliases: dict[str, str],
) -> dict | None:
    if not registered_name:
        return None
    if registered_version is not None:
        version_key = f"{registered_name}:{registered_version}"
        if version_key in signature_map:
            return signature_map[version_key]["payload"]
        alias = signature_aliases.get(version_key)
        if alias and alias in signature_map:
            return signature_map[alias]["payload"]
    if registered_name in signature_map:
        return signature_map[registered_name]["payload"]
    alias = signature_aliases.get(registered_name)
    if alias and alias in signature_map:
        return signature_map[alias]["payload"]

    registered_norm = normalize_model_name(registered_name)
    matches = []
    for key in signature_map:
        key_name = key.split(":", 1)[0]
        key_norm = normalize_model_name(key_name)
        if not registered_norm or not key_norm:
            continue
        if registered_norm == key_norm or registered_norm in key_norm or key_norm in registered_norm:
            matches.append(key)
    if len(matches) == 1:
        return signature_map[matches[0]]["payload"]
    return None


def get_mlflow_model_version_source(
    tracking_uri: str, model_name: str, model_version: int
) -> str | None:
    url = f"{tracking_uri.rstrip('/')}/api/2.0/mlflow/model-versions/get"
    resp = requests.get(
        url, params={"name": model_name, "version": model_version}, timeout=15
    )
    resp.raise_for_status()
    payload = resp.json()
    model_version_info = payload.get("model_version", {})
    source = model_version_info.get("source")
    if isinstance(source, str) and source:
        return source
    return None


def load_mlflow_input_example(artifact_uri: str) -> dict | None:
    try:
        from mlflow import artifacts
    except Exception:
        return None
    os.environ.setdefault("MLFLOW_ENABLE_ARTIFACTS_PROGRESS_BAR", "false")
    local_dir = artifacts.download_artifacts(artifact_uri=artifact_uri)
    for name in ("serving_input_example.json", "input_example.json"):
        path = Path(local_dir) / name
        if path.exists():
            return json.loads(path.read_text())
    return None


def safe_get(label: str, func):
    try:
        return func(), None
    except requests.RequestException as exc:
        return None, f"{label}: {exc}"


def resolve_domino_url() -> tuple[str, str]:
    candidates = [
        DOMINO_URL_ENV,
        DOMINO_USER_HOST_ENV,
        DOMINO_API_HOST_ENV,
        DOMINO_API_PROXY_ENV,
    ]
    for name in candidates:
        value = read_env(name)
        if value:
            return value, name
    return DEFAULT_DOMINO_URL, DOMINO_URL_ENV


def looks_internal_host(url: str) -> bool:
    return any(token in url for token in ("domino-platform", "localhost", "127.0.0.1"))


def detect_public_host(domino_url: str, api_key: str) -> str:
    if api_key:
        try:
            url = f"{domino_url.rstrip('/')}/v4/auth/principal"
            resp = requests.get(url, headers={"X-Domino-Api-Key": api_key}, timeout=5)
            resp.raise_for_status()
            payload = resp.json()
            apps_host = payload.get("appsHost")
            apps_subdomain = payload.get("appsSubdomain")
            if isinstance(apps_host, str) and apps_host:
                if isinstance(apps_subdomain, str) and apps_subdomain:
                    match = re.match(r"^(https?://)(.+)$", apps_host)
                    if match:
                        scheme, host = match.groups()
                        if host.startswith(apps_subdomain):
                            return f"{scheme}{host[len(apps_subdomain):]}"
                return apps_host
        except (requests.RequestException, ValueError, json.JSONDecodeError):
            pass

    try:
        resp = requests.get(domino_url, timeout=5)
    except requests.RequestException:
        return ""
    match = re.search(r"https?://[^\"']+/auth/realms/", resp.text)
    if match:
        return match.group(0).split("/auth/realms/")[0]
    match = re.search(r"https?://[^\"']+/auth/", resp.text)
    if match:
        return match.group(0).split("/auth/")[0]
    return ""


def resolve_inference_url(domino_url: str, domino_url_env: str, api_key: str) -> tuple[str, str]:
    candidates = [
        DOMINO_INFERENCE_BASE_URL_ENV,
        DOMINO_INFERENCE_HOST_ENV,
        DOMINO_PUBLIC_HOST_ENV,
    ]
    for name in candidates:
        value = read_env(name)
        if value:
            return value, name
    if domino_url and looks_internal_host(domino_url):
        detected = detect_public_host(domino_url, api_key)
        if detected:
            return detected, ""
    return domino_url, domino_url_env


def ensure_https_port(url: str) -> str:
    parsed = urlparse(url)
    if parsed.scheme != "https":
        return url
    if parsed.netloc and ":" in parsed.netloc:
        return url
    netloc = f"{parsed.netloc}:443" if parsed.netloc else parsed.netloc
    return urlunparse(parsed._replace(netloc=netloc))


def build_inference_url(base_url: str, model_id: str) -> str:
    return f"{base_url.rstrip('/')}/models/{model_id}/latest/model"


def fetch_model_overview_html(domino_url: str, model_id: str, api_key: str) -> str:
    url = f"{domino_url.rstrip('/')}/models/{model_id}/overview"
    headers = {"X-Domino-Api-Key": api_key}
    resp = requests.get(url, headers=headers, timeout=15)
    resp.raise_for_status()
    return resp.text


def extract_overview_curl(html_text: str) -> str | None:
    match = re.search(
        r'<div role="tabpanel" class="tab-pane" id="language-curl">.*?<pre[^>]*>(.*?)</pre>',
        html_text,
        flags=re.S | re.I,
    )
    if not match:
        return None
    curl_text = html.unescape(match.group(1))
    return curl_text.strip()


def replace_curl_data(curl_text: str, payload: dict) -> str:
    payload_json = json.dumps(payload)
    replacement = f"-d '{payload_json}'"
    lines = curl_text.splitlines()
    for idx, line in enumerate(lines):
        if re.search(r"(^|\s)-d\s+['\"]", line):
            indent = re.match(r"\s*", line).group(0)
            suffix_match = re.search(r"\s*\\\s*$", line)
            suffix = suffix_match.group(0) if suffix_match else ""
            lines[idx] = f"{indent}{replacement}{suffix}"
            return "\n".join(lines)
    insert_at = len(lines)
    for idx, line in enumerate(lines):
        if re.search(r"\s-u\s+", line):
            insert_at = idx
            break
    lines.insert(insert_at, replacement)
    return "\n".join(lines)


def main() -> None:
    domino_url, domino_url_env = resolve_domino_url()
    api_key = read_env(DOMINO_API_KEY_ENV)
    project_id = read_env(DOMINO_PROJECT_ID_ENV)
    require_env(domino_url, api_key, project_id)
    inference_url, inference_url_env = resolve_inference_url(domino_url, domino_url_env, api_key)
    inference_url = ensure_https_port(inference_url)

    v4_base = f"{domino_url}/v4"
    signature_map = extract_local_signature()
    signature_aliases = parse_signature_aliases()
    tracking_uri = read_env(MLFLOW_TRACKING_URI_ENV)
    mlflow_payloads: dict[tuple[str, int], dict | None] = {}

    errors = []
    model_manager_models, err = safe_get(
        "modelManager.getModels",
        lambda: get_json(
            build_url(v4_base, "/modelManager/getModels", {"projectId": project_id}),
            api_key,
        ),
    )
    if err:
        errors.append(err)

    model_products, err = safe_get(
        "modelProducts.list",
        lambda: get_json(
            build_url(v4_base, "/modelProducts", {"projectId": project_id}),
            api_key,
        ),
    )
    if err:
        errors.append(err)

    model_apis, err = safe_get(
        "modelServing.listModelApis",
        lambda: get_json(
            build_url(domino_url, "/api/modelServing/v1/modelApis", {"projectId": project_id}),
            api_key,
        ),
    )
    if err:
        errors.append(err)

    genai_endpoints, err = safe_get(
        "genAI.endpoints",
        lambda: get_json(
            build_url(domino_url, "/api/gen-ai/beta/endpoints", {"projectId": project_id}),
            api_key,
        ),
    )
    if err:
        errors.append(err)

    registered_models, err = safe_get(
        "registeredModels.list",
        lambda: get_json(
            build_url(domino_url, "/api/registeredmodels/v1", {"projectId": project_id}),
            api_key,
        ),
    )
    if err:
        errors.append(err)

    curls = []
    inference_entries = []

    list_curls = [
        build_curl(
            "GET",
            build_url(v4_base, "/modelManager/getModels", {"projectId": project_id}),
            api_key,
        ),
        build_curl(
            "GET",
            build_url(v4_base, "/modelProducts", {"projectId": project_id}),
            api_key,
        ),
        build_curl(
            "GET",
            build_url(domino_url, "/api/modelServing/v1/modelApis", {"projectId": project_id}),
            api_key,
        ),
        build_curl(
            "GET",
            build_url(domino_url, "/api/gen-ai/beta/endpoints", {"projectId": project_id}),
            api_key,
        ),
        build_curl(
            "GET",
            build_url(domino_url, "/api/registeredmodels/v1", {"projectId": project_id}),
            api_key,
        ),
    ]
    curls.append({"label": "Project listings", "entries": list_curls})

    if isinstance(model_manager_models, list):
        entries = []
        for model in model_manager_models:
            model_id = model.get("id")
            name = model.get("name")
            active = model.get("activeVersion") or {}
            version_id = active.get("id")
            signature = None
            file_info = active.get("file")
            if isinstance(file_info, dict):
                file_info = file_info.get("path")
            function = active.get("function")
            registered_name = active.get("registeredModelName")
            registered_version = active.get("registeredModelVersion")
            if file_info or function:
                signature = f"{file_info}:{function}".strip(":")
            payload = resolve_signature_payload(
                registered_name, registered_version, signature_map, signature_aliases
            )
            model_payload = signature_to_model_payload(payload)
            mlflow_payload = None
            if (
                tracking_uri
                and isinstance(registered_name, str)
                and registered_name
                and isinstance(registered_version, int)
            ):
                key = (registered_name, registered_version)
                if key not in mlflow_payloads:
                    try:
                        source = get_mlflow_model_version_source(
                            tracking_uri, registered_name, registered_version
                        )
                        if source:
                            example = load_mlflow_input_example(source)
                            mlflow_payloads[key] = example_to_model_payload(example)
                        else:
                            mlflow_payloads[key] = None
                    except requests.RequestException as exc:
                        errors.append(f"mlflow.modelVersion[{registered_name}:{registered_version}]: {exc}")
                        mlflow_payloads[key] = None
                    except (OSError, ValueError, json.JSONDecodeError) as exc:
                        errors.append(
                            f"mlflow.modelVersion[{registered_name}:{registered_version}]: {exc}"
                        )
                        mlflow_payloads[key] = None
                mlflow_payload = mlflow_payloads.get(key)

            entry = {
                "name": name,
                "id": model_id,
                "activeVersionId": version_id,
                "signature": signature,
            }

            entry_curls = []
            if model_id and version_id:
                entry_curls.append(
                    build_curl(
                        "GET",
                        build_url(
                            v4_base,
                            f"/modelManager/{model_id}/{version_id}/usage",
                            {"limit": 25},
                        ),
                        api_key,
                    )
                )
                entry_curls.append(
                    build_curl(
                        "GET",
                        build_url(
                            v4_base,
                            f"/modelManager/{model_id}/{version_id}/httpRequests",
                            {"limit": 25},
                        ),
                        api_key,
                    )
                )

            if model.get("isAsync") and model_id:
                async_url = build_url(domino_url, f"/api/modelApis/async/v1/{model_id}")
                entry_curls.append(
                    build_curl("POST", async_url, api_key, payload or {"parameters": {}})
                )
            elif payload and model_id:
                entry_curls.append(
                    build_curl(
                        "POST",
                        build_url(domino_url, f"/api/modelApis/async/v1/{model_id}"),
                        api_key,
                        payload,
                    )
                )

            entry["curls"] = entry_curls
            entries.append(entry)

            if model_id:
                overview_curl = None
                overview_html, err = safe_get(
                    f"modelOverview[{model_id}]",
                    lambda: fetch_model_overview_html(domino_url, model_id, api_key),
                )
                if err:
                    errors.append(err)
                elif overview_html:
                    overview_curl = extract_overview_curl(overview_html)

                effective_payload = mlflow_payload or model_payload
                payload = effective_payload or {"data": {"default": 345.3, "income": 245}}
                if overview_curl and effective_payload:
                    overview_curl = replace_curl_data(overview_curl, effective_payload)
                inference_entries.append(
                    {
                        "name": name,
                        "id": model_id,
                        "payload": payload,
                        "overview_curl": overview_curl,
                    }
                )
        curls.append({"label": "Model Manager models", "entries": entries})

    if isinstance(model_products, list):
        entries = []
        for product in model_products:
            model_product_id = product.get("id")
            name = product.get("name")
            open_url = product.get("openUrl") or product.get("runningAppUrl")
            entry = {"name": name, "id": model_product_id}
            entry_curls = []
            if model_product_id:
                entry_curls.append(
                    build_curl(
                        "GET",
                        build_url(v4_base, f"/modelProducts/{model_product_id}"),
                        api_key,
                    )
                )
            if open_url:
                entry_curls.append(build_curl("GET", open_url, api_key))
            entry["curls"] = entry_curls
            entries.append(entry)
        curls.append({"label": "Model products", "entries": entries})

    if isinstance(model_apis, dict):
        entries = []
        for api in model_apis.get("items", []):
            api_id = api.get("id")
            name = api.get("name")
            entry = {"name": name, "id": api_id}
            entry_curls = []
            if api_id:
                entry_curls.append(
                    build_curl(
                        "GET",
                        build_url(domino_url, f"/api/modelServing/v1/modelApis/{api_id}"),
                        api_key,
                    )
                )
                entry_curls.append(
                    build_curl(
                        "GET",
                        build_url(domino_url, f"/api/modelServing/v1/modelApis/{api_id}/versions"),
                        api_key,
                    )
                )
            entry["curls"] = entry_curls
            entries.append(entry)
        curls.append({"label": "Model Serving APIs", "entries": entries})

    if isinstance(genai_endpoints, dict):
        entries = []
        for endpoint in genai_endpoints.get("items", []):
            endpoint_id = endpoint.get("id")
            name = endpoint.get("name")
            model_source = None
            current_version = endpoint.get("currentVersion") or {}
            model_source = current_version.get("modelSource") or {}
            registered = model_source.get("registeredModel") or {}
            registered_name = registered.get("modelName")
            payload = None
            if registered_name and registered_name in signature_map:
                payload = signature_map[registered_name]["payload"]

            entry_curls = []
            url = endpoint.get("vanityUrl") or endpoint.get("url")
            if url and payload:
                entry_curls.append(build_curl("POST", url, api_key, payload))
            elif url:
                entry_curls.append(build_curl("POST", url, api_key, {"input": "TODO"}))

            if endpoint_id:
                entry_curls.append(
                    build_curl(
                        "GET",
                        build_url(domino_url, f"/api/gen-ai/beta/endpoints/{endpoint_id}"),
                        api_key,
                    )
                )
            entries.append({"name": name, "id": endpoint_id, "curls": entry_curls})
        curls.append({"label": "Gen AI endpoints", "entries": entries})

    if isinstance(registered_models, dict):
        entries = []
        for model in registered_models.get("items", []):
            name = model.get("name")
            latest_version = model.get("latestVersion")
            entry_curls = []
            if name:
                entry_curls.append(
                    build_curl(
                        "GET",
                        build_url(domino_url, f"/api/registeredmodels/v1/{name}"),
                        api_key,
                    )
                )
                entry_curls.append(
                    build_curl(
                        "GET",
                        build_url(domino_url, f"/api/registeredmodels/v1/{name}/versions"),
                        api_key,
                    )
                )
                if latest_version is not None:
                    entry_curls.append(
                        build_curl(
                            "GET",
                            build_url(
                                domino_url,
                                f"/api/registeredmodels/v1/{name}/versions/{latest_version}",
                            ),
                            api_key,
                        )
                    )
            entries.append({"name": name, "latestVersion": latest_version, "curls": entry_curls})
        curls.append({"label": "Registered models", "entries": entries})

    for section in curls:
        print(f"== {section['label']} ==")
        for item in section["entries"]:
            if isinstance(item, str):
                print(item)
                continue
            header = item.get("name") or item.get("label") or item.get("id") or "item"
            print(f"- {header}")
            signature = item.get("signature")
            if signature:
                print(f"  signature: {signature}")
            for curl in item.get("curls", []):
                print(f"  {curl}")
        print()

    if signature_map:
        print("== Local model signatures ==")
        for name, info in signature_map.items():
            payload = json.dumps(info["payload"])
            print(f"- {name} ({info['source']}): {payload}")
        print()

    if errors:
        print("== Errors ==")
        for err in errors:
            print(f"- {err}")

    if isinstance(model_manager_models, list):
        print("== Model inference curls ==")
        for item in inference_entries:
            header = item.get("name") or item.get("id") or "model"
            print(f"- {header}")
            model_id = item.get("id")
            overview_curl = item.get("overview_curl")
            if overview_curl:
                print(overview_curl)
                continue
            url = build_inference_url(inference_url, model_id)
            print(f"  {build_inference_curl(url, api_key, item['payload'])}")
        print()


if __name__ == "__main__":
    main()
