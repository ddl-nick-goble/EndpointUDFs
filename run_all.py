#!/usr/bin/env python3
"""
Combined Excel-DNA Add-in Generator for Domino Model Endpoints

This script:
1. Fetches all models for a given Domino project
2. Extracts curl commands and credentials from the model overview pages
3. Retrieves parameter signatures from MLflow (if available)
4. Generates a complete Excel-DNA add-in (.xll) with User Defined Functions (UDFs)

Each discovered endpoint is converted into a strongly-typed UDF with full documentation,
parameter descriptions, and error handling.
"""

import base64
import html
import json
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from typing import Any

import requests

# Environment configuration
DOMINO_URL = os.environ.get("DOMINO_URL", "https://se-demo.domino.tech:443")
API_KEY = os.environ.get("DOMINO_USER_API_KEY", "")
PROJECT_ID = os.environ.get("DOMINO_PROJECT_ID", "")
PROJECT_NAME = os.environ.get("DOMINO_PROJECT_NAME", "")


@dataclass
class EndpointConfig:
    """Configuration for a single API endpoint to be converted to a UDF."""
    name: str
    url: str
    username: str
    password: str
    parameters: list[dict[str, Any]]  # List of {name, type, description, example}
    description: str
    return_description: str


# =============================================================================
# Endpoint Discovery Functions (from claude_create_curls.py)
# =============================================================================

def get_models(project_id: str) -> list:
    """Get all models for a project."""
    url = f"{DOMINO_URL}/v4/modelManager/getModels"
    headers = {"X-Domino-Api-Key": API_KEY}
    resp = requests.get(url, params={"projectId": project_id}, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json()


def _load_input_example(local_dir: str) -> dict | None:
    """Load the input example from MLflow artifacts, if present."""
    for fname in ("serving_input_example.json", "input_example.json"):
        path = os.path.join(local_dir, fname)
        if os.path.exists(path):
            with open(path) as f:
                example = json.load(f)
            # Convert dataframe_split format to simple dict
            if "dataframe_split" in example:
                cols = example["dataframe_split"]["columns"]
                data_rows = example["dataframe_split"]["data"]
                if not data_rows:
                    return {"data": {col: None for col in cols}}
                if len(data_rows) == 1:
                    row = data_rows[0]
                    return {"data": {col: row[i] for i, col in enumerate(cols)}}
                col_data = {col: [] for col in cols}
                for row in data_rows:
                    for i, col in enumerate(cols):
                        col_data[col].append(row[i])
                return {"data": col_data}
            if "data" in example:
                return example
    return None


def _load_signature_inputs(local_dir: str) -> list[dict[str, Any]] | None:
    """Load the MLflow model signature inputs from the MLmodel file."""
    mlmodel_path = os.path.join(local_dir, "MLmodel")
    if not os.path.exists(mlmodel_path):
        return None

    try:
        import yaml
    except Exception:
        return None

    with open(mlmodel_path) as f:
        mlmodel = yaml.safe_load(f)

    signature = mlmodel.get("signature") if isinstance(mlmodel, dict) else None
    if not signature:
        return None

    inputs = signature.get("inputs")
    if not inputs:
        return None

    if isinstance(inputs, str):
        try:
            inputs = json.loads(inputs)
        except json.JSONDecodeError:
            return None

    if isinstance(inputs, dict) and "inputs" in inputs:
        inputs = inputs["inputs"]

    if not isinstance(inputs, list):
        return None

    return inputs


def get_model_signature(model_name: str, model_version: int) -> dict | None:
    """Get the signature and input example for a registered model version from MLflow."""
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

    # Download the model artifacts and parse signature + input example
    try:
        from mlflow import artifacts
        os.environ.setdefault("MLFLOW_ENABLE_ARTIFACTS_PROGRESS_BAR", "false")
        local_dir = artifacts.download_artifacts(artifact_uri=source)

        signature_inputs = _load_signature_inputs(local_dir)
        example = _load_input_example(local_dir)
        if signature_inputs or example:
            return {"signature_inputs": signature_inputs, "example": example}
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


def parse_curl_command(curl_text: str) -> dict | None:
    """
    Parse a curl command to extract URL, credentials, and payload.

    Returns dict with keys: url, username, password, data
    """
    if not curl_text:
        return None

    result = {}

    # Extract URL - look for the URL in the curl command
    # Usually follows 'curl' and comes before or after flags
    url_match = re.search(r'https?://[^\s\'"]+', curl_text)
    if url_match:
        result['url'] = url_match.group(0)

    # Extract basic auth credentials (-u or --user flag)
    # Format: -u username:password or --user username:password
    auth_match = re.search(r'(?:-u|--user)\s+[\'"]?([^:]+):([^\s\'"]+)[\'"]?', curl_text)
    if auth_match:
        result['username'] = auth_match.group(1)
        result['password'] = auth_match.group(2)

    # Extract the data payload (-d flag)
    # Could be -d 'JSON' or -d "JSON"
    data_match = re.search(r"-d\s+'([^']*)'", curl_text)
    if not data_match:
        data_match = re.search(r'-d\s+"([^"]*)"', curl_text)
    if data_match:
        try:
            result['data'] = json.loads(data_match.group(1))
        except json.JSONDecodeError:
            result['data'] = None

    return result if 'url' in result else None


def extract_help_topic_url(endpoint_url: str, domino_base_url: str) -> str:
    """
    Generate Domino model overview URL for HelpTopic.

    Extracts model_id from endpoint URL and builds a link to the model overview page.
    Example: https://se-demo.domino.tech/models/696a8cffde2f14747436d217/latest/model
             → https://se-demo.domino.tech/models/696a8cffde2f14747436d217/overview?ownerName=nick_goble&projectName=EndpointUDFs
    """
    # Extract model_id from endpoint URL (handles /api/ or /latest/model patterns)
    match = re.search(r'/models/([a-f0-9]+)/', endpoint_url)
    if not match:
        return ""

    model_id = match.group(1)
    # Use the base URL from DOMINO_URL env var, strip port if present
    base_url = domino_base_url.rstrip(':443').rstrip('/')

    # Build overview URL with owner and project
    # Note: These could be parameterized in the future if needed
    return f"{base_url}/models/{model_id}/overview?ownerName=nick_goble&projectName=EndpointUDFs"


def _split_param_tokens(name: str) -> list[str]:
    """Split a parameter name into lowercase tokens for heuristics."""
    if not name:
        return []
    tokens = re.findall(r"[A-Z]?[a-z]+|[0-9]+|[A-Z]+(?![a-z])", name)
    if not tokens:
        tokens = re.split(r"[^a-zA-Z0-9]+", name)
    return [t.lower() for t in tokens if t]


def _looks_like_date_string(value: str) -> bool:
    """Detect common date formats like yyyy-mm-dd or ISO timestamps."""
    if not value:
        return False
    return bool(re.match(r"^\d{4}[-/]\d{2}[-/]\d{2}([ T].*)?$", value.strip()))


def _is_date_param(name: str, value: Any) -> bool:
    """Heuristic detection for date-like parameters."""
    tokens = _split_param_tokens(name)
    if any(t in {"date", "dt", "dob"} for t in tokens):
        return True
    if isinstance(value, str) and _looks_like_date_string(value):
        return True
    return False


def _map_mlflow_type(type_name: str, name: str, example: Any) -> str:
    """Map MLflow type name to a UDF parameter type."""
    if not type_name:
        return "string" if _is_date_param(name, example) else infer_parameter_type(example)

    lower = type_name.lower()
    if lower in {"boolean", "bool"}:
        return "bool"
    if lower in {"integer", "long", "int", "short"}:
        return "double"
    if lower in {"double", "float", "float32", "float64"}:
        return "double"
    if lower in {"date", "datetime"}:
        return "date"
    if lower in {"string", "str"}:
        return "date" if _is_date_param(name, example) else "string"
    return "object"


def _parse_mlflow_input_type(spec: dict[str, Any]) -> tuple[str, bool]:
    """Parse MLflow input spec to get base type and array flag."""
    type_name = str(spec.get("type", "") or "")

    if type_name.lower() == "tensor":
        tensor_spec = spec.get("tensor-spec", {}) or {}
        dtype = str(tensor_spec.get("dtype", "") or "")
        return _map_mlflow_type(dtype, spec.get("name", ""), None), True

    if type_name.lower() == "array":
        items = spec.get("items") or {}
        if isinstance(items, dict):
            item_type = str(items.get("type", "") or "")
        else:
            item_type = str(items)
        return _map_mlflow_type(item_type, spec.get("name", ""), None), True

    array_match = re.match(r"array\s*<\s*([^>]+)\s*>", type_name, flags=re.I)
    if not array_match:
        array_match = re.match(r"array\s*\(\s*([^)]+)\s*\)", type_name, flags=re.I)
    if not array_match:
        array_match = re.match(r"array\s*\[\s*([^\]]+)\s*\]", type_name, flags=re.I)
    if array_match:
        base_type = array_match.group(1)
        return _map_mlflow_type(base_type, spec.get("name", ""), None), True

    return _map_mlflow_type(type_name, spec.get("name", ""), None), False


def infer_parameter_type(value: Any) -> str:
    """Infer the C# type from a Python value."""
    if isinstance(value, bool):
        return "bool"
    elif isinstance(value, int):
        return "double"  # Use double for all numbers in Excel
    elif isinstance(value, float):
        return "double"
    elif isinstance(value, str):
        return "string"
    else:
        return "object"


def _should_allow_reference(param_type: str) -> bool:
    """
    Determine if a parameter should allow cell references.
    Numeric and date params can accept cell references for better UX.
    """
    return param_type in {"double", "date", "number"}


def _infer_array_element_type(value: Any) -> str:
    """Infer the element type for list/tuple values, handling nested lists."""
    if isinstance(value, (list, tuple)):
        for item in value:
            if isinstance(item, (list, tuple)):
                nested_type = _infer_array_element_type(item)
                if nested_type != "object":
                    return nested_type
            else:
                return infer_parameter_type(item)
    return "object"


def clean_function_name(name: str, prefix: str = "Model") -> str:
    """
    Clean a model name for use as a function name.
    Only applies CamelCase transformation if there's punctuation/spaces to remove.

    Examples:
        'hedging-model' -> 'HedgingModel'
        'my_cool_model' -> 'MyCoolModel'
        'HedgingModel' -> 'HedgingModel' (unchanged)
        'SimpleModel' -> 'SimpleModel' (unchanged)
    """
    # If the name is already clean (only alphanumeric), return as-is
    if re.match(r'^[a-zA-Z][a-zA-Z0-9]*$', name):
        return name

    # Otherwise, split on non-alphanumeric and CamelCase it
    parts = re.split(r'[^a-zA-Z0-9]+', name)
    camel = ''.join(part.capitalize() for part in parts if part)

    # Ensure it starts with a letter
    if camel and camel[0].isdigit():
        camel = prefix + camel
    return camel or f'Unnamed{prefix}'


def get_project_name(project_id: str) -> str:
    """Get the project name from environment variable or Domino API."""
    if PROJECT_NAME:
        return clean_function_name(PROJECT_NAME, prefix="Project")

    url = f"{DOMINO_URL}/v4/projects/{project_id}"
    headers = {"X-Domino-Api-Key": API_KEY}
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        name = resp.json().get("name", "")
        if name:
            return clean_function_name(name, prefix="Project")
    except Exception:
        pass

    return ""


def discover_endpoints(project_id: str, project_name: str) -> list[EndpointConfig]:
    """
    Discover all model endpoints in a project and build EndpointConfig objects.
    """
    endpoints = []
    models = get_models(project_id)

    for model in models:
        model_id = model.get("id")
        name = model.get("name", "UnnamedModel")
        active = model.get("activeVersion") or {}
        registered_name = active.get("registeredModelName")
        registered_version = active.get("registeredModelVersion")

        if not model_id:
            continue

        print(f"  Discovering: {name}...")

        # Get the curl from the overview page
        curl_text = get_curl_from_html(model_id)
        if not curl_text:
            print(f"    (skipped - no curl found)")
            continue

        # Parse the curl command
        curl_info = parse_curl_command(curl_text)
        if not curl_info or 'username' not in curl_info or 'password' not in curl_info:
            print(f"    (skipped - could not parse curl)")
            continue

        # Try to get the signature from MLflow
        signature = None
        if registered_name and registered_version:
            signature = get_model_signature(registered_name, registered_version)

        signature_inputs = None
        example_data = None
        if signature:
            signature_inputs = signature.get("signature_inputs")
            example = signature.get("example")
            if example and isinstance(example, dict):
                example_data = example.get("data")

        # If no MLflow signature, try to use the data from the curl command
        if not signature_inputs and curl_info.get('data'):
            signature_inputs = None
            example_data = curl_info.get("data", {}).get("data")

        # Build parameters from the MLflow signature (preferred)
        parameters = []
        if signature_inputs:
            for spec in signature_inputs:
                param_name = spec.get("name")
                if not param_name:
                    continue
                example_value = None
                if example_data and param_name in example_data:
                    example_value = example_data[param_name]
                param_type, is_array = _parse_mlflow_input_type(spec)
                if _is_date_param(param_name, example_value) and param_type == "string":
                    param_type = "date"
                parameters.append({
                    "name": param_name,
                    "type": param_type,
                    "is_array": is_array,
                    "description": f"The {param_name} parameter for the model",
                    "example": example_value
                })

        # Fallback: derive parameters from example data if signature is missing
        if not parameters and isinstance(example_data, dict):
            for param_name, param_value in example_data.items():
                is_array = isinstance(param_value, (list, tuple))
                if is_array:
                    param_type = _infer_array_element_type(param_value)
                    if param_type == "object":
                        param_type = "string"
                else:
                    param_type = infer_parameter_type(param_value)
                if _is_date_param(param_name, param_value) and param_type == "string":
                    param_type = "date"
                parameters.append({
                    "name": param_name,
                    "type": param_type,
                    "is_array": is_array,
                    "description": f"The {param_name} parameter for the model",
                    "example": param_value
                })

        if not parameters:
            print(f"    (skipped - no parameters found in signature)")
            continue

        # Create the endpoint config
        # Use the endpoint name, cleaned up only if it has punctuation
        function_name = clean_function_name(name)
        endpoint = EndpointConfig(
            name=function_name,
            url=curl_info['url'],
            username=curl_info['username'],
            password=curl_info['password'],
            parameters=parameters,
            description=f"Calls the {name} Domino model API endpoint.",
            return_description="Returns the model result value (spills across cells if array)"
        )
        endpoints.append(endpoint)
        param_names = ", ".join([p["name"] for p in parameters])
        excel_function_name = f"Domino.{project_name}.{function_name}" if project_name else f"Domino.{function_name}"
        print(f"    Found: {excel_function_name}({param_names})")

    return endpoints


# =============================================================================
# Code Generation Functions (from claude_create_udfs.py)
# =============================================================================

def generate_udf_method(endpoint: EndpointConfig, project_name: str) -> str:
    """Generate a C# UDF method for a single endpoint."""

    # Build ExcelArgument attributes and parameter section
    param_section_parts = []
    for p in endpoint.parameters:
        # Allow reference for array inputs or numeric/date params.
        allow_ref_str = "true" if p.get("is_array") or _should_allow_reference(p["type"]) else "false"
        excel_arg = f'[ExcelArgument(Name = "{p["name"]}", Description = "{p["description"]}", AllowReference = {allow_ref_str})]'
        param_type = "object"
        param_section_parts.append(f'{excel_arg} {param_type} {p["name"]}')

    param_section = ", ".join(param_section_parts)

    # Build JSON payload construction based on parameter types
    # We generate C# code that builds a proper JSON string with concatenation
    # Target output example for strings: "{\"data\": {\"key\": \"" + val.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"}}"
    # Target output example for numbers: "{\"data\": {\"key\": " + val.ToString(...) + "}}"
    json_parts = []
    for i, p in enumerate(endpoint.parameters):
        param_name = p["name"]
        is_last = (i == len(endpoint.parameters) - 1)
        separator = "" if is_last else ", "
        param_kind = p["type"]
        if param_kind == "object":
            param_kind = "string"
        if param_kind not in {"string", "bool", "date"}:
            param_kind = "number"
        kind_enum = {
            "string": "ParamKind.String",
            "bool": "ParamKind.Bool",
            "date": "ParamKind.Date",
            "number": "ParamKind.Number",
        }[param_kind]
        allow_array = "true" if p.get("is_array") else "false"
        json_parts.append(
            f'"\\\"{param_name}\\\": " + SerializeParamValue({param_name}, {kind_enum}, {allow_array}) + "{separator}"'
        )

    # Join the parts with + operator
    json_inner = " + ".join(json_parts)
    json_construction = f'"{{\\\"data\\\": {{" + {json_inner} + "}}}}"'

    # Base64 encode credentials
    credentials = f'{endpoint.username}:{endpoint.password}'
    auth_header = base64.b64encode(credentials.encode()).decode()

    # Escape description for C# string
    escaped_description = endpoint.description.replace('"', '\\"')

    # Excel function name with Domino. prefix
    if project_name:
        excel_function_name = f"Domino.{project_name}.{endpoint.name}"
    else:
        excel_function_name = f"Domino.{endpoint.name}"

    # Generate HelpTopic URL if possible
    help_topic_url = extract_help_topic_url(endpoint.url, DOMINO_URL)
    help_topic_line = f',\n            HelpTopic = "{help_topic_url}"' if help_topic_url else ''

    method = f'''
        /// <summary>
        /// {endpoint.description}
        /// </summary>
        /// <returns>{endpoint.return_description}</returns>
        [ExcelFunction(
            Name = "{excel_function_name}",
            Description = "{escaped_description}",
            Category = "Domino Model APIs",
            IsVolatile = false,
            IsExceptionSafe = true,
            IsThreadSafe = true{help_topic_line}
        )]
        public static object {endpoint.name}(
            {param_section})
        {{
            try
            {{
                // Force TLS 1.2 (required for modern HTTPS endpoints)
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

                string url = "{endpoint.url}";
                string jsonPayload = {json_construction};

                using (var client = new WebClient())
                {{
                    client.Headers[HttpRequestHeader.ContentType] = "application/json";
                    client.Headers[HttpRequestHeader.Authorization] = "Basic {auth_header}";

                    string response = client.UploadString(url, "POST", jsonPayload);
                    return ParseResult(response);
                }}
            }}
            catch (WebException ex)
            {{
                if (ex.Response != null)
                {{
                    using (var reader = new StreamReader(ex.Response.GetResponseStream()))
                    {{
                        return "API Error: " + reader.ReadToEnd();
                    }}
                }}
                return "Error: " + ex.Message;
            }}
            catch (Exception ex)
            {{
                return "Error: " + ex.Message;
            }}
        }}
'''
    return method


def generate_csharp_code(endpoints: list[EndpointConfig], project_name: str) -> str:
    """Generate the complete C# add-in code."""

    methods = "\n".join([generate_udf_method(ep, project_name) for ep in endpoints])

    # Build function documentation
    func_docs = "\n".join([
        f"/// - Domino.{project_name}.{ep.name}: {ep.description[:60]}..."
        if project_name else f"/// - Domino.{ep.name}: {ep.description[:60]}..."
        for ep in endpoints
    ])

    code = f'''using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

/// <summary>
/// Excel-DNA Add-in providing UDFs for Domino Model API endpoints.
///
/// This add-in was auto-generated and provides the following functions:
{func_docs}
///
/// Each function calls a specific Domino model endpoint with the provided parameters
/// and returns the result from the model (supports array spilling for multiple results).
/// </summary>
public static class DominoModelFunctions
{{
    private enum ParamKind
    {{
        String,
        Number,
        Bool,
        Date,
    }}

    /// <summary>
    /// Formats a date-like parameter into yyyy-MM-dd.
    /// Accepts Excel dates, Unix epoch (seconds/ms), and date strings.
    /// </summary>
    private static string FormatDateParam(object value)
    {{
        if (value == null)
        {{
            return "";
        }}
        if (value is ExcelMissing || value is ExcelEmpty)
        {{
            return "";
        }}

        if (value is double d)
        {{
            return FormatDateFromNumber(d);
        }}
        if (value is int i)
        {{
            return FormatDateFromNumber(i);
        }}
        if (value is DateTime dt)
        {{
            return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }}
        if (value is string s)
        {{
            s = s.Trim();
            if (string.IsNullOrEmpty(s))
            {{
                return "";
            }}

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
            {{
                return FormatDateFromNumber(num);
            }}

            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime parsed))
            {{
                return parsed.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }}

            return s;
        }}

        return value.ToString();
    }}

    private static string FormatDateFromNumber(double value)
    {{
        // Epoch milliseconds or seconds
        if (value >= 1_000_000_000_000d)
        {{
            var dt = DateTimeOffset.FromUnixTimeMilliseconds((long)Math.Round(value)).DateTime;
            return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }}
        if (value >= 1_000_000_000d)
        {{
            var dt = DateTimeOffset.FromUnixTimeSeconds((long)Math.Round(value)).DateTime;
            return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }}

        try
        {{
            var dt = DateTime.FromOADate(value);
            return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }}
        catch
        {{
            return value.ToString(CultureInfo.InvariantCulture);
        }}
    }}

    private static string EscapeJsonString(string value)
    {{
        if (value == null)
        {{
            return "";
        }}
        return value.Replace("\\\\", "\\\\\\\\").Replace("\\"", "\\\\\\"");
    }}

    private static object NormalizeExcelValue(object value)
    {{
        if (value is ExcelReference excelRef)
        {{
            try
            {{
                value = excelRef.GetValue();
            }}
            catch
            {{
                return value;
            }}
        }}
        return value;
    }}

    private static string SerializeParamValue(object value, ParamKind kind, bool allowArray)
    {{
        value = NormalizeExcelValue(value);
        if (value == null || value is ExcelMissing || value is ExcelEmpty)
        {{
            return "null";
        }}

        if (allowArray && value is Array array)
        {{
            return SerializeArrayValue(array, kind);
        }}

        if (!allowArray && value is Array singleCell && singleCell.Rank == 2 && singleCell.GetLength(0) == 1 && singleCell.GetLength(1) == 1)
        {{
            return SerializeScalarValue(singleCell.GetValue(0, 0), kind);
        }}

        return SerializeScalarValue(value, kind);
    }}

    private static bool IsEmptyCell(object value)
    {{
        return value == null || value is ExcelMissing || value is ExcelEmpty || value is ExcelError;
    }}

    private static string SerializeArrayValue(Array array, ParamKind kind)
    {{
        if (array.Rank == 1)
        {{
            var sb = new StringBuilder();
            sb.Append('[');
            int start = array.GetLowerBound(0);
            int end = array.GetUpperBound(0);
            bool appended = false;
            for (int i = start; i <= end; i++)
            {{
                object value = array.GetValue(i);
                if (IsEmptyCell(value))
                {{
                    continue;
                }}
                if (appended)
                {{
                    sb.Append(',');
                }}
                sb.Append(kind == ParamKind.Number ? SerializeScalarValue(Convert.ToString(value, CultureInfo.InvariantCulture), ParamKind.String) : SerializeScalarValue(value, kind));
                appended = true;
            }}
            sb.Append(']');
            return sb.ToString();
        }}

        if (array.Rank == 2)
        {{
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);
            var sb = new StringBuilder();
            sb.Append('[');

            if (rows == 1 || cols == 1)
            {{
                int count = rows == 1 ? cols : rows;
                bool appended = false;
                for (int i = 0; i < count; i++)
                {{
                    object value = rows == 1 ? array.GetValue(0, i) : array.GetValue(i, 0);
                    if (IsEmptyCell(value))
                    {{
                        continue;
                    }}
                    if (appended)
                    {{
                        sb.Append(',');
                    }}
                    sb.Append(kind == ParamKind.Number ? SerializeScalarValue(Convert.ToString(value, CultureInfo.InvariantCulture), ParamKind.String) : SerializeScalarValue(value, kind));
                    appended = true;
                }}
                sb.Append(']');
                return sb.ToString();
            }}

            for (int r = 0; r < rows; r++)
            {{
                if (r > 0)
                {{
                    sb.Append(',');
                }}
                sb.Append('[');
                bool appended = false;
                for (int c = 0; c < cols; c++)
                {{
                    object value = array.GetValue(r, c);
                    if (IsEmptyCell(value))
                    {{
                        continue;
                    }}
                    if (appended)
                    {{
                        sb.Append(',');
                    }}
                    sb.Append(kind == ParamKind.Number ? SerializeScalarValue(Convert.ToString(value, CultureInfo.InvariantCulture), ParamKind.String) : SerializeScalarValue(value, kind));
                    appended = true;
                }}
                sb.Append(']');
            }}
            sb.Append(']');
            return sb.ToString();
        }}

        return SerializeScalarValue(array, kind);
    }}

    private static string SerializeScalarValue(object value, ParamKind kind)
    {{
        if (value == null || value is ExcelMissing || value is ExcelEmpty)
        {{
            return "null";
        }}

        switch (kind)
        {{
            case ParamKind.String:
                return "\\"" + EscapeJsonString(Convert.ToString(value, CultureInfo.InvariantCulture)) + "\\"";
            case ParamKind.Date:
                return "\\"" + EscapeJsonString(FormatDateParam(value)) + "\\"";
            case ParamKind.Bool:
                if (value is bool b)
                {{
                    return b ? "true" : "false";
                }}
                if (value is double d)
                {{
                    return d != 0d ? "true" : "false";
                }}
                if (value is int i)
                {{
                    return i != 0 ? "true" : "false";
                }}
                if (value is string s)
                {{
                    if (bool.TryParse(s, out bool parsedBool))
                    {{
                        return parsedBool ? "true" : "false";
                    }}
                    if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedNum))
                    {{
                        return parsedNum != 0d ? "true" : "false";
                    }}
                }}
                return "false";
            default:
                return SerializeNumberValue(value, false);
        }}
    }}

    private static string SerializeNumberValue(object value, bool forceFloat)
    {{
        if (value is double num)
        {{
            return FormatNumber(num, forceFloat);
        }}
        if (value is int numInt)
        {{
            return FormatNumber(numInt, forceFloat);
        }}
        if (value is string str)
        {{
            if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed))
            {{
                return FormatNumber(parsed, forceFloat);
            }}
            return "\\"" + EscapeJsonString(str) + "\\"";
        }}
        try
        {{
            return FormatNumber(Convert.ToDouble(value, CultureInfo.InvariantCulture), forceFloat);
        }}
        catch
        {{
            return "\\"" + EscapeJsonString(Convert.ToString(value, CultureInfo.InvariantCulture)) + "\\"";
        }}
    }}

    private static string FormatNumber(double value, bool forceFloat)
    {{
        if (!forceFloat)
        {{
            return value.ToString(CultureInfo.InvariantCulture);
        }}
        if (Math.Abs(value % 1) < 1e-12)
        {{
            return value.ToString("0.0", CultureInfo.InvariantCulture);
        }}
        return value.ToString(CultureInfo.InvariantCulture);
    }}

    /// <summary>
    /// Extracts the "result" field from the JSON response and returns it as Excel-friendly output.
    /// Handles single values, 1D arrays (horizontal spill), and 2D arrays (grid spill).
    /// </summary>
    private static object ParseResult(string json)
    {{
        if (!TryExtractResultValue(json, out string resultValue, out string error))
        {{
            return error;
        }}

        // Check if it's a 2D array (array of arrays) like [[1,2],[3,4]]
        if (resultValue.StartsWith("[["))
        {{
            return Parse2DArray(resultValue);
        }}
        // Check if it's a 1D array like [1,2,3]
        else if (resultValue.StartsWith("[") && resultValue.EndsWith("]"))
        {{
            return Parse1DArray(resultValue);
        }}
        else
        {{
            // Single value
            return ParseSingleValue(resultValue);
        }}
    }}

    /// <summary>
    /// Extracts the raw JSON value for the "result" field without relying on regex for nested arrays.
    /// </summary>
    private static bool TryExtractResultValue(string json, out string resultValue, out string error)
    {{
        resultValue = "";
        error = "";

        var match = Regex.Match(json, @"""result""\s*:");
        if (!match.Success)
        {{
            error = "Error: No result field in response";
            return false;
        }}

        int i = match.Index + match.Length;
        while (i < json.Length && char.IsWhiteSpace(json[i]))
        {{
            i++;
        }}

        if (i >= json.Length)
        {{
            error = "Error: Empty result field in response";
            return false;
        }}

        char start = json[i];
        if (start == '[' || start == '{{')
        {{
            char open = start;
            char close = (start == '[') ? ']' : '}}';
            int depth = 0;
            bool inString = false;
            bool escape = false;
            int startIndex = i;

            for (; i < json.Length; i++)
            {{
                char ch = json[i];
                if (inString)
                {{
                    if (escape)
                    {{
                        escape = false;
                        continue;
                    }}
                    if (ch == '\\\\')
                    {{
                        escape = true;
                        continue;
                    }}
                    if (ch == '"')
                    {{
                        inString = false;
                    }}
                    continue;
                }}

                if (ch == '"')
                {{
                    inString = true;
                    continue;
                }}

                if (ch == open)
                {{
                    depth++;
                }}
                else if (ch == close)
                {{
                    depth--;
                    if (depth == 0)
                    {{
                        resultValue = json.Substring(startIndex, i - startIndex + 1).Trim();
                        return true;
                    }}
                }}
            }}

            error = "Error: Unterminated result value in response";
            return false;
        }}

        if (start == '"')
        {{
            int startIndex = i;
            bool escape = false;
            for (i = i + 1; i < json.Length; i++)
            {{
                char ch = json[i];
                if (escape)
                {{
                    escape = false;
                    continue;
                }}
                if (ch == '\\\\')
                {{
                    escape = true;
                    continue;
                }}
                if (ch == '"')
                {{
                    resultValue = json.Substring(startIndex, i - startIndex + 1).Trim();
                    return true;
                }}
            }}
            error = "Error: Unterminated string result in response";
            return false;
        }}

        int primitiveStart = i;
        while (i < json.Length && json[i] != ',' && json[i] != '}}' && json[i] != ']')
        {{
            i++;
        }}
        resultValue = json.Substring(primitiveStart, i - primitiveStart).Trim();
        return true;
    }}

    /// <summary>
    /// Parses a 2D array like [[1,2,3],[4,5,6]] into an Excel-compatible object[,] for grid spill.
    /// </summary>
    private static object Parse2DArray(string arrayStr)
    {{
        // Extract inner arrays using regex to find each [...] row
        var rowMatches = Regex.Matches(arrayStr, @"\[([^\[\]]*)\]");
        if (rowMatches.Count == 0)
        {{
            return "Error: Invalid 2D array format";
        }}

        // Parse each row to get dimensions and values
        var rows = new List<List<object>>();
        int maxCols = 0;

        foreach (Match rowMatch in rowMatches)
        {{
            string rowContent = rowMatch.Groups[1].Value;
            var rowValues = new List<object>();

            if (!string.IsNullOrWhiteSpace(rowContent))
            {{
                var parts = rowContent.Split(new[] {{ ',' }}, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                {{
                    rowValues.Add(ParseSingleValue(part.Trim()));
                }}
            }}

            rows.Add(rowValues);
            if (rowValues.Count > maxCols)
            {{
                maxCols = rowValues.Count;
            }}
        }}

        // Handle edge cases
        if (rows.Count == 0 || maxCols == 0)
        {{
            return "";
        }}
        if (rows.Count == 1 && rows[0].Count == 1)
        {{
            return rows[0][0];
        }}

        // Create 2D array for Excel grid spill
        object[,] spillArray = new object[rows.Count, maxCols];
        for (int r = 0; r < rows.Count; r++)
        {{
            for (int c = 0; c < maxCols; c++)
            {{
                if (c < rows[r].Count)
                {{
                    spillArray[r, c] = rows[r][c];
                }}
                else
                {{
                    spillArray[r, c] = ""; // Pad jagged arrays
                }}
            }}
        }}
        return spillArray;
    }}

    /// <summary>
    /// Parses a 1D array like [1,2,3] into an Excel-compatible object[,] for horizontal spill.
    /// </summary>
    private static object Parse1DArray(string arrayStr)
    {{
        string inner = arrayStr.Substring(1, arrayStr.Length - 2).Trim();
        if (string.IsNullOrEmpty(inner))
        {{
            return "";
        }}

        var parts = inner.Split(new[] {{ ',' }}, StringSplitOptions.RemoveEmptyEntries);
        var results = new List<object>();

        foreach (var part in parts)
        {{
            results.Add(ParseSingleValue(part.Trim()));
        }}

        if (results.Count == 1)
        {{
            return results[0];
        }}

        // Create a 1-row, N-column array for horizontal spill
        object[,] spillArray = new object[1, results.Count];
        for (int i = 0; i < results.Count; i++)
        {{
            spillArray[0, i] = results[i];
        }}
        return spillArray;
    }}

    /// <summary>
    /// Parses a single value (number or string) into the appropriate type.
    /// </summary>
    private static object ParseSingleValue(string value)
    {{
        string trimmed = value.Trim();
        if (double.TryParse(trimmed, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out double numVal))
        {{
            return numVal;
        }}
        return trimmed.Trim('"');
    }}

{methods}
}}
'''
    return code


def generate_dna_file(project_name: str) -> str:
    """Generate the Excel-DNA .dna configuration file."""
    addin_name = "Domino endpoint UDFs for project"
    if project_name:
        addin_name = f"{addin_name} - {project_name}"
    return '''<?xml version="1.0" encoding="utf-8"?>
<DnaLibrary Name="''' + addin_name + '''" RuntimeVersion="v4.0">
  <ExternalLibrary Path="DominoModelFunctions.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" />
</DnaLibrary>
'''


def build_addin(endpoints: list[EndpointConfig], project_name: str) -> str | None:
    """Build the Excel-DNA add-in."""

    if not endpoints:
        print("No endpoints to build. Exiting.")
        return None

    print()
    print("=" * 60)
    print("Building Excel Add-in")
    print("=" * 60)
    print()

    # Create a temporary build directory
    build_dir = tempfile.mkdtemp(prefix="exceldna_build_")
    print(f"[1/6] Created temporary build directory: {build_dir}")

    try:
        # Write the C# code
        cs_file = os.path.join(build_dir, "DominoModelFunctions.cs")
        with open(cs_file, "w") as f:
            f.write(generate_csharp_code(endpoints, project_name))
        print(f"[2/6] Generated C# source code with {len(endpoints)} UDF(s):")
        for ep in endpoints:
            params = ", ".join([p["name"] for p in ep.parameters])
            function_name = f"Domino.{project_name}.{ep.name}" if project_name else f"Domino.{ep.name}"
            print(f"       - {function_name}({params})")

        # Write the .dna file
        dna_file = os.path.join(build_dir, "DominoModelFunctions.dna")
        with open(dna_file, "w") as f:
            f.write(generate_dna_file(project_name))
        print("[3/6] Generated Excel-DNA configuration file")

        # Create a .csproj file for building
        csproj_content = '''<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <OutputType>Library</OutputType>
    <GenerateAssemblyInfo>false</GenerateAssemblyInfo>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="ExcelDna.AddIn" Version="1.7.0" />
  </ItemGroup>
</Project>
'''
        csproj_file = os.path.join(build_dir, "DominoModelFunctions.csproj")
        with open(csproj_file, "w") as f:
            f.write(csproj_content)
        print("[4/6] Generated project file")

        # Run dotnet restore and build
        print("[5/6] Building add-in (this may take a moment)...")

        # Restore packages
        result = subprocess.run(
            ["dotnet", "restore"],
            cwd=build_dir,
            capture_output=True,
            text=True
        )
        if result.returncode != 0:
            print(f"       Restore output: {result.stdout}")
            print(f"       Restore errors: {result.stderr}")
            raise RuntimeError(f"dotnet restore failed: {result.stderr}")

        # Build the project
        result = subprocess.run(
            ["dotnet", "build", "-c", "Release"],
            cwd=build_dir,
            capture_output=True,
            text=True
        )
        if result.returncode != 0:
            print(f"       Build output: {result.stdout}")
            print(f"       Build errors: {result.stderr}")
            raise RuntimeError(f"dotnet build failed: {result.stderr}")

        print("       Build completed successfully!")

        # Find the PACKED .xll files
        publish_dir = os.path.join(build_dir, "bin", "Release", "net48", "publish")

        src_xll_64 = None
        src_xll_32 = None

        if os.path.exists(publish_dir):
            for f in os.listdir(publish_dir):
                if f.endswith("-packed.xll"):
                    full_path = os.path.join(publish_dir, f)
                    if "64" in f:
                        src_xll_64 = full_path
                    else:
                        src_xll_32 = full_path

        if not src_xll_64 and not src_xll_32:
            # Fallback: search entire build directory for packed xll files
            for root, dirs, files in os.walk(build_dir):
                for f in files:
                    if f.endswith("-packed.xll"):
                        full_path = os.path.join(root, f)
                        if "64" in f:
                            src_xll_64 = full_path
                        else:
                            src_xll_32 = full_path

        if not src_xll_64 and not src_xll_32:
            raise RuntimeError("Could not find packed .xll files. Check build output.")

        # Copy to current directory
        copied_files = []
        if src_xll_64:
            dest_xll_64 = os.path.join(os.getcwd(), "DominoModelFunctions-AddIn64.xll")
            shutil.copy(src_xll_64, dest_xll_64)
            copied_files.append(("64-bit", dest_xll_64))

        if src_xll_32:
            dest_xll_32 = os.path.join(os.getcwd(), "DominoModelFunctions-AddIn.xll")
            shutil.copy(src_xll_32, dest_xll_32)
            copied_files.append(("32-bit", dest_xll_32))

        print(f"[6/6] Add-in created successfully!")
        for arch, path in copied_files:
            print(f"       {arch}: {path}")

        # Write source files for reference
        cs_output = os.path.join(os.getcwd(), "DominoModelFunctions.cs")
        dna_output = os.path.join(os.getcwd(), "DominoModelFunctions.dna")
        with open(cs_output, "w") as f:
            f.write(generate_csharp_code(endpoints, project_name))
        with open(dna_output, "w") as f:
            f.write(generate_dna_file(project_name))

        print()
        print("=" * 60)
        print("SUCCESS! Your Excel add-in is ready.")
        print("=" * 60)
        print()
        print("To use the add-in:")
        print("  1. Open Excel")
        print("  2. Go to File > Options > Add-ins")
        print("  3. At the bottom, select 'Excel Add-ins' and click 'Go...'")
        print("  4. Click 'Browse...' and select the .xll file")
        print("  5. Click OK to enable the add-in")
        print()
        print("Available functions:")
        for ep in endpoints:
            params = ", ".join([p["name"] for p in ep.parameters])
            function_name = f"Domino.{project_name}.{ep.name}" if project_name else f"Domino.{ep.name}"
            print(f"  ={function_name}({params})")
            print(f"    {ep.description[:70]}...")
            print()

        artifacts_dir = "/mnt/artifacts"
        os.makedirs(artifacts_dir, exist_ok=True)
        source_64 = "/mnt/code/DominoModelFunctions-AddIn64.xll"
        source_32 = "/mnt/code/DominoModelFunctions-AddIn.xll"
        dest_64 = os.path.join(artifacts_dir, "DominoExcelUDFsAddIn64.xll")
        dest_32 = os.path.join(artifacts_dir, "DominoExcelUDFsAddIn.xll")
        if os.path.exists(source_64):
            shutil.copy(source_64, dest_64)
        else:
            print(f"Warning: missing source file {source_64}")
        if os.path.exists(source_32):
            shutil.copy(source_32, dest_32)
        else:
            print(f"Warning: missing source file {source_32}")

        return copied_files[0][1] if copied_files else None

    finally:
        # Clean up the temp directory
        shutil.rmtree(build_dir, ignore_errors=True)


def main():
    """Main entry point - discover endpoints and build add-in."""

    print("=" * 60)
    print("Domino Model APIs - Combined Endpoint Discovery & Add-in Generator")
    print("=" * 60)
    print()

    # Check required environment variables
    if not API_KEY:
        print("Error: DOMINO_USER_API_KEY environment variable is required")
        return

    project_id = PROJECT_ID
    if not project_id:
        print("Error: DOMINO_PROJECT_ID environment variable is required")
        return

    print(f"Domino URL: {DOMINO_URL}")
    print(f"Project ID: {project_id}")
    print()

    project_name = get_project_name(project_id)

    # Step 1: Discover endpoints
    print("Step 1: Discovering model endpoints...")
    print("-" * 40)
    endpoints = discover_endpoints(project_id, project_name)

    if not endpoints:
        print()
        print("No valid endpoints discovered. Check that:")
        print("  - The project has deployed models with active versions")
        print("  - The models have input signatures (from MLflow or curl data)")
        print("  - Your API key has access to view the models")
        return

    print()
    print(f"Discovered {len(endpoints)} endpoint(s)")
    print()

    # Step 2: Build the add-in
    print("Step 2: Building Excel add-in...")
    print("-" * 40)

    try:
        build_addin(endpoints, project_name)
    except Exception as e:
        print(f"\nBuild Error: {e}")
        print("\nTroubleshooting:")
        print("  - Ensure .NET SDK (6.0+) is installed: dotnet --version")
        print("  - On Linux/Mac, you may need to build on Windows for .xll generation")
        print()

        # Write fallback artifacts to disk for manual builds
        cs_output = os.path.join(os.getcwd(), "DominoModelFunctions.cs")
        dna_output = os.path.join(os.getcwd(), "DominoModelFunctions.dna")
        with open(cs_output, "w") as f:
            f.write(generate_csharp_code(endpoints, project_name))
        with open(dna_output, "w") as f:
            f.write(generate_dna_file(project_name))
        print(f"Fallback files written:")
        print(f"  - {cs_output}")
        print(f"  - {dna_output}")


if __name__ == "__main__":
    main()
