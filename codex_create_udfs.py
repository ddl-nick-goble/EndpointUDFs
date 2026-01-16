import json
import os
import re
import shutil
import subprocess
import time
import zipfile
from pathlib import Path

import requests

CURL_SPECS = [
    {
        "name": "FirstTestEndpoint",
        "curl": r"""curl \
'https://se-demo.domino.tech:443/models/69694398de2f14747436d1ae/latest/model' \
-H 'Content-Type: application/json' \
-d '{"data": {"age": 35.0, "income": 85000.0}}'  \
-u b34AehOxIAtA2qjasXbucudBmZTIi9z4wEWwungZaTwpujigDF5EyEj6ayAZzVw1:b34AehOxIAtA2qjasXbucudBmZTIi9z4wEWwungZaTwpujigDF5EyEj6ayAZzVw1
""",
    },
    {
        "name": "SecondTestEndpoint",
        "curl": r"""curl \
'https://se-demo.domino.tech:443/models/69696646de2f14747436d1b3/latest/model' \
-H 'Content-Type: application/json' \
-d '{"data": {"age": 35.05, "income": 95000.1}}'  \
-u mFNhCmCMGlqvw32S2We0LSXByp1k03AsNkd5EPP4QsSllQ8r2LCwHZS2PGkTW2Kl:mFNhCmCMGlqvw32S2We0LSXByp1k03AsNkd5EPP4QsSllQ8r2LCwHZS2PGkTW2Kl
""",
    },
]

NUGET_EXCELDNA_ADDIN_URL = "https://www.nuget.org/api/v2/package/ExcelDna.AddIn/1.7.0"


def parse_curl(curl_text: str) -> dict:
    url_match = re.search(r"https?://[^'\"\\s]+", curl_text)
    auth_match = re.search(r"-u\\s+([^\\s]+)", curl_text)
    data_match = re.search(r"-d\\s+'([^']*)'", curl_text) or re.search(r"-d\\s+\"([^\"]*)\"", curl_text)
    url = url_match.group(0) if url_match else ""
    auth = auth_match.group(1) if auth_match else ""
    data_raw = data_match.group(1) if data_match else ""
    payload = parse_payload(data_raw)
    user, password = ("", "")
    if auth and ":" in auth:
        user, password = auth.split(":", 1)
    return {
        "url": url,
        "auth_user": user,
        "auth_pass": password,
        "payload": payload,
    }


def parse_payload(raw: str) -> dict | None:
    if not raw:
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        start = raw.find("{")
        end = raw.rfind("}")
        if start != -1 and end != -1 and end > start:
            try:
                return json.loads(raw[start : end + 1])
            except json.JSONDecodeError:
                return None
    return None


def payload_to_params(payload: dict | None) -> list[str]:
    if not isinstance(payload, dict):
        return []
    data = payload.get("data")
    if not isinstance(data, dict):
        return []
    return [str(key) for key in data.keys()]


def sanitize_identifier(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]+", "_", name)
    if re.match(r"^[0-9]", cleaned):
        cleaned = f"Udf_{cleaned}"
    return cleaned.strip("_") or "Udf"


def build_csharp(udfs: list[dict]) -> str:
    lines = [
        "using System;",
        "using System.Collections.Generic;",
        "using System.Net.Http;",
        "using System.Net.Http.Headers;",
        "using System.Text;",
        "using System.Text.Json;",
        "using ExcelDna.Integration;",
        "",
        "namespace DominoUdfs",
        "{",
        "    public static class DominoUdfs",
        "    {",
        "        private static readonly HttpClient Client = new HttpClient",
        "        {",
        "            Timeout = TimeSpan.FromSeconds(30)",
        "        };",
        "",
        "        private static string CallEndpoint(string url, string user, string pass, object payload)",
        "        {",
        "            try",
        "            {",
        "                var json = JsonSerializer.Serialize(payload);",
        "                using var request = new HttpRequestMessage(HttpMethod.Post, url);",
        "                var token = Convert.ToBase64String(Encoding.UTF8.GetBytes($\"{user}:{pass}\"));",
        "                request.Headers.Authorization = new AuthenticationHeaderValue(\"Basic\", token);",
        "                request.Content = new StringContent(json, Encoding.UTF8, \"application/json\");",
        "                using var response = Client.SendAsync(request).GetAwaiter().GetResult();",
        "                var body = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();",
        "                return $\"{(int)response.StatusCode} {response.StatusCode}: {body}\";",
        "            }",
        "            catch (Exception ex)",
        "            {",
        "                return $\"ERROR: {ex.Message}\";",
        "            }",
        "        }",
        "",
    ]

    for udf in udfs:
        name = udf["name"]
        identifier = sanitize_identifier(name)
        params = udf["params"]
        url = udf["url"]
        user = udf["auth_user"]
        password = udf["auth_pass"]
        description = f"Calls Domino endpoint {name} and returns status + response body."
        args = []
        for param in params:
            args.append(
                f'[ExcelArgument(Name="{param}", Description="Input for `{param}` from the model signature.")]'
                f" double {sanitize_identifier(param)}"
            )
        args_text = ", ".join(args) if args else ""
        data_pairs = ", ".join(
            [f'{{\"{param}\", {sanitize_identifier(param)}}}' for param in params]
        )
        lines.extend(
            [
                "        [ExcelFunction(",
                f'            Name = "{identifier}",',
                f'            Description = "{description}"',
                "        )]",
                f"        public static string {identifier}({args_text})",
                "        {",
                f"            var payload = new Dictionary<string, object> {{",
                f'                {{"data", new Dictionary<string, object> {{ {data_pairs} }} }}',
                "            };",
                f'            return CallEndpoint("{url}", "{user}", "{password}", payload);',
                "        }",
                "",
            ]
        )

    lines.extend(["    }", "}"])
    return "\n".join(lines)


def build_csproj() -> str:
    return "\n".join(
        [
            '<Project Sdk="Microsoft.NET.Sdk">',
            "  <PropertyGroup>",
            "    <TargetFramework>net6.0</TargetFramework>",
            "    <Nullable>enable</Nullable>",
            "    <LangVersion>latest</LangVersion>",
            "  </PropertyGroup>",
            "  <ItemGroup>",
            '    <PackageReference Include="ExcelDna.Integration" Version="1.7.0" />',
            "  </ItemGroup>",
            "</Project>",
        ]
    )


def build_dna(addin_name: str, dll_name: str) -> str:
    return "\n".join(
        [
            f'<DnaLibrary Name="{addin_name}" RuntimeVersion="v4.0">',
            f'  <ExternalLibrary Path="{dll_name}" />',
            "</DnaLibrary>",
        ]
    )


def download_exceldna_addin(to_dir: Path) -> Path:
    to_dir.mkdir(parents=True, exist_ok=True)
    archive_path = to_dir / "ExcelDna.AddIn.nupkg"
    resp = requests.get(NUGET_EXCELDNA_ADDIN_URL, timeout=60)
    resp.raise_for_status()
    archive_path.write_bytes(resp.content)
    with zipfile.ZipFile(archive_path, "r") as zf:
        zf.extractall(to_dir)
    return to_dir


def main() -> None:
    workdir = Path.cwd()
    build_root = workdir / ".codex_udf_build"
    build_root.mkdir(parents=True, exist_ok=True)

    udfs = []
    for spec in CURL_SPECS:
        parsed = parse_curl(spec["curl"])
        params = payload_to_params(parsed.get("payload"))
        udfs.append(
            {
                "name": spec["name"],
                "url": parsed["url"],
                "auth_user": parsed["auth_user"],
                "auth_pass": parsed["auth_pass"],
                "params": params,
            }
        )

    csproj_dir = build_root / "DominoUdfs"
    csproj_dir.mkdir(parents=True, exist_ok=True)
    (csproj_dir / "DominoUdfs.csproj").write_text(build_csproj())
    (csproj_dir / "DominoUdfs.cs").write_text(build_csharp(udfs))

    dotnet = shutil.which("dotnet")
    if not dotnet:
        raise SystemExit("dotnet is required to build the ExcelDna add-in.")

    subprocess.run(
        [dotnet, "build", "-c", "Release"],
        check=True,
        cwd=csproj_dir,
    )

    dll_path = csproj_dir / "bin" / "Release" / "net6.0" / "DominoUdfs.dll"
    if not dll_path.exists():
        raise SystemExit(f"Missing build output: {dll_path}")

    stamp = time.strftime("%Y%m%d_%H%M%S")
    output_dir = workdir / "excel_addins"
    output_dir.mkdir(parents=True, exist_ok=True)
    addin_base = f"DominoUdfs_{stamp}"
    dna_path = output_dir / f"{addin_base}.dna"
    xll_path = output_dir / f"{addin_base}.xll"
    dll_out = output_dir / "DominoUdfs.dll"

    dll_out.write_bytes(dll_path.read_bytes())
    dna_path.write_text(build_dna(addin_base, dll_out.name))

    exceldna_dir = download_exceldna_addin(build_root / "exceldna")
    xll_source = exceldna_dir / "tools" / "net6.0-windows7.0" / "ExcelDna64.xll"
    if not xll_source.exists():
        xll_source = exceldna_dir / "tools" / "net452" / "ExcelDna64.xll"
    if not xll_source.exists():
        raise SystemExit("ExcelDna64.xll not found in ExcelDna.AddIn package.")
    xll_path.write_bytes(xll_source.read_bytes())

    pack_exe = exceldna_dir / "tools" / "ExcelDnaPack.exe"
    mono = shutil.which("mono")
    if pack_exe.exists() and mono:
        subprocess.run([mono, str(pack_exe), str(dna_path), "/Y"], check=True, cwd=output_dir)
        packed = output_dir / f"{addin_base}-packed.xll"
        if packed.exists():
            xll_path.unlink(missing_ok=True)
            packed.rename(xll_path)

    print(f"Add-in created: {xll_path}")
    print(f"Companion files: {dna_path}, {dll_out}")


if __name__ == "__main__":
    main()
