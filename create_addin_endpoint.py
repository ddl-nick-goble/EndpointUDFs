#!/usr/bin/env python3
import argparse
import json
import os
import sys
from typing import Any

import requests

DOMINO_URL = os.environ.get("DOMINO_URL", "https://se-demo.domino.tech:443")
API_KEY = os.environ.get("DOMINO_USER_API_KEY", "")
PROJECT_ID = os.environ.get("DOMINO_PROJECT_ID", "")
DEFAULT_NAME = "ExcelWebAddinHosting"


def _require(value: str, label: str) -> str:
    if not value:
        print(f"Missing required {label}.", file=sys.stderr)
        sys.exit(2)
    return value


def _parse_env_vars(values: list[str]) -> list[dict[str, str]]:
    env_vars: list[dict[str, str]] = []
    for item in values:
        if "=" not in item:
            print(f"Invalid --env '{item}'. Use KEY=VALUE.", file=sys.stderr)
            sys.exit(2)
        key, value = item.split("=", 1)
        env_vars.append({"key": key, "value": value})
    return env_vars


def _request(method: str, path: str, *, params: dict[str, Any] | None = None,
             payload: dict[str, Any] | None = None) -> Any:
    url = f"{DOMINO_URL}{path}"
    headers = {"X-Domino-Api-Key": API_KEY}
    resp = requests.request(
        method,
        url,
        headers=headers,
        params=params,
        json=payload,
        timeout=30,
    )
    if resp.status_code >= 400:
        print(f"Request failed: {resp.status_code} {resp.text}", file=sys.stderr)
        resp.raise_for_status()
    if resp.text:
        return resp.json()
    return None


def find_model_api(project_id: str, name: str) -> dict[str, Any] | None:
    params = {"projectId": project_id, "name": name, "limit": 25, "offset": 0}
    data = _request("GET", "/api/modelServing/v1/modelApis", params=params)
    if not data:
        return None
    items = data.get("items", [])
    for item in items:
        if item.get("name") == name and not item.get("isArchived", False):
            return item
    return None


def build_payload(args: argparse.Namespace, project_id: str) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "name": args.name,
        "description": args.description,
        "environmentId": args.environment_id,
        "environmentVariables": _parse_env_vars(args.env),
        "isAsync": args.is_async,
        "strictNodeAntiAffinity": args.strict_node_anti_affinity,
        "version": {
            "projectId": project_id,
            "logHttpRequestResponse": args.log_http_request_response,
            "monitoringEnabled": args.monitoring_enabled,
            "source": {},
        },
    }

    if args.replicas is not None:
        payload["replicas"] = args.replicas
    if args.hardware_tier_id:
        payload["hardwareTierId"] = args.hardware_tier_id
    if args.resource_quota_id:
        payload["resourceQuotaId"] = args.resource_quota_id
    if args.bundle_id:
        payload["bundleId"] = args.bundle_id

    version: dict[str, Any] = payload["version"]
    if args.version_description:
        version["description"] = args.version_description
    if args.commit_id:
        version["commitId"] = args.commit_id
    if args.environment_id:
        version["environmentId"] = args.environment_id
    if args.provenance_checkpoint_id:
        version["provenanceCheckpointId"] = args.provenance_checkpoint_id
    if args.prediction_dataset_resource_id:
        version["predictionDatasetResourceId"] = args.prediction_dataset_resource_id
    if args.record_invocation is not None:
        version["recordInvocation"] = args.record_invocation
    if args.should_deploy is not None:
        version["shouldDeploy"] = args.should_deploy

    if args.source_type == "Registry":
        _require(args.registered_model_name, "registered model name")
        _require(args.registered_model_version, "registered model version")
        version["source"] = {
            "type": "Registry",
            "registeredModelName": args.registered_model_name,
            "registeredModelVersion": int(args.registered_model_version),
        }
    else:
        version["source"] = {
            "type": "File",
            "file": args.model_file,
            "function": args.model_function,
        }

    return payload


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create ExcelWebAddinHosting Model API if missing.",
    )
    parser.add_argument("--project-id", default=PROJECT_ID)
    parser.add_argument("--name", default=DEFAULT_NAME)
    parser.add_argument("--description", default="Office Add-in static host")
    parser.add_argument("--environment-id", default=os.environ.get("DOMINO_MODEL_API_ENVIRONMENT_ID", ""))
    parser.add_argument("--bundle-id", default=os.environ.get("DOMINO_MODEL_API_BUNDLE_ID", ""))
    parser.add_argument("--hardware-tier-id", default=os.environ.get("DOMINO_MODEL_API_HARDWARE_TIER_ID", ""))
    parser.add_argument("--resource-quota-id", default=os.environ.get("DOMINO_MODEL_API_RESOURCE_QUOTA_ID", ""))
    parser.add_argument("--replicas", type=int, default=1)
    parser.add_argument("--is-async", action="store_true", default=False)
    parser.add_argument("--strict-node-anti-affinity", action="store_true", default=False)
    parser.add_argument("--env", action="append", default=[], help="Environment variable KEY=VALUE")

    parser.add_argument("--version-description", default="")
    parser.add_argument("--commit-id", default=os.environ.get("DOMINO_MODEL_API_COMMIT_ID", ""))
    parser.add_argument("--provenance-checkpoint-id",
                        default=os.environ.get("DOMINO_MODEL_API_PROVENANCE_CHECKPOINT_ID", ""))
    parser.add_argument("--prediction-dataset-resource-id",
                        default=os.environ.get("DOMINO_MODEL_API_PREDICTION_DATASET_RESOURCE_ID", ""))
    parser.add_argument("--log-http-request-response", action="store_true", default=False)
    parser.add_argument("--monitoring-enabled", action="store_true", default=False)
    parser.add_argument("--record-invocation", action="store_true", default=False)
    parser.add_argument("--should-deploy", dest="should_deploy", action="store_true")
    parser.add_argument("--no-should-deploy", dest="should_deploy", action="store_false")
    parser.set_defaults(should_deploy=True)

    parser.add_argument("--source-type", choices=["File", "Registry"], default="File")
    parser.add_argument("--model-file", default="model.py")
    parser.add_argument("--model-function", default="predict")
    parser.add_argument("--registered-model-name", default="")
    parser.add_argument("--registered-model-version", default="")

    args = parser.parse_args()

    _require(API_KEY, "DOMINO_USER_API_KEY")
    project_id = _require(args.project_id, "project id")
    _require(args.environment_id, "environment id (set --environment-id or DOMINO_MODEL_API_ENVIRONMENT_ID)")

    existing = find_model_api(project_id, args.name)
    if existing:
        model_id = existing.get("id")
        print(f"Model API already exists: {args.name} ({model_id})")
        return

    payload = build_payload(args, project_id)
    created = _request("POST", "/api/modelServing/v1/modelApis", payload=payload)
    model_id = created.get("id") if isinstance(created, dict) else None
    print(f"Created Model API: {args.name} ({model_id})")
    print(json.dumps(created, indent=2))


if __name__ == "__main__":
    main()
