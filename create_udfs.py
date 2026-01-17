#!/usr/bin/env python3
"""
Excel-DNA Add-in Generator for Domino Model Endpoints

This script generates a complete Excel-DNA add-in (.xll) with User Defined Functions (UDFs)
for calling Domino model API endpoints directly from Excel.

Each curl endpoint is converted into a strongly-typed UDF with full documentation,
parameter descriptions, and error handling.
"""

import os
import subprocess
import shutil
import tempfile
from dataclasses import dataclass
from typing import List, Dict, Any
import base64


@dataclass
class EndpointConfig:
    """Configuration for a single API endpoint to be converted to a UDF."""
    name: str
    url: str
    username: str
    password: str
    parameters: List[Dict[str, Any]]  # List of {name, type, description, example}
    description: str
    return_description: str


# Define the endpoints extracted from the curl commands
ENDPOINTS = [
    EndpointConfig(
        name="FirstTestEndpoint",
        url="https://se-demo.domino.tech:443/models/69694398de2f14747436d1ae/latest/model",
        username="b34AehOxIAtA2qjasXbucudBmZTIi9z4wEWwungZaTwpujigDF5EyEj6ayAZzVw1",
        password="b34AehOxIAtA2qjasXbucudBmZTIi9z4wEWwungZaTwpujigDF5EyEj6ayAZzVw1",
        parameters=[
            {
                "name": "age",
                "type": "double",
                "description": "The age value for the prediction model (numeric)",
                "example": 35.0
            },
            {
                "name": "income",
                "type": "double",
                "description": "The income value for the prediction model (numeric)",
                "example": 85000.0
            }
        ],
        description="Calls the FirstTestEndpoint Domino model API with age and income parameters. "
                    "This endpoint appears to be a prediction model that takes demographic data.",
        return_description="Returns the model result value (spills across cells if array)"
    ),
    EndpointConfig(
        name="SecondTestEndpoint",
        url="https://se-demo.domino.tech:443/models/69696646de2f14747436d1b3/latest/model",
        username="mFNhCmCMGlqvw32S2We0LSXByp1k03AsNkd5EPP4QsSllQ8r2LCwHZS2PGkTW2Kl",
        password="mFNhCmCMGlqvw32S2We0LSXByp1k03AsNkd5EPP4QsSllQ8r2LCwHZS2PGkTW2Kl",
        parameters=[
            {
                "name": "age",
                "type": "double",
                "description": "The age value for the prediction model (numeric, supports decimals)",
                "example": 35.05
            },
            {
                "name": "income",
                "type": "double",
                "description": "The income value for the prediction model (numeric, supports decimals)",
                "example": 95000.1
            }
        ],
        description="Calls the SecondTestEndpoint Domino model API with age and income parameters. "
                    "This endpoint appears to be a similar prediction model with potentially different underlying logic.",
        return_description="Returns the model result value (spills across cells if array)"
    )
]


def generate_udf_method(endpoint: EndpointConfig) -> str:
    """Generate a C# UDF method for a single endpoint."""

    # Build parameter list for method signature
    param_declarations = ", ".join([
        f'{p["type"]} {p["name"]}' for p in endpoint.parameters
    ])

    # Build ExcelArgument attributes
    excel_args = []
    for p in endpoint.parameters:
        excel_args.append(
            f'[ExcelArgument(Name = "{p["name"]}", Description = "{p["description"]}")]'
        )

    # Build the full parameter section with attributes
    param_section = ", ".join([
        f'{excel_args[i]} {endpoint.parameters[i]["type"]} {endpoint.parameters[i]["name"]}'
        for i in range(len(endpoint.parameters))
    ])

    # Build JSON payload construction
    json_parts = ", ".join([
        f'\\"{p["name"]}\\": " + {p["name"]}.ToString(System.Globalization.CultureInfo.InvariantCulture) + "'
        for p in endpoint.parameters
    ])
    json_construction = f'"{{\\\"data\\\": {{{json_parts}}}}}"'

    # Base64 encode credentials
    credentials = f'{endpoint.username}:{endpoint.password}'
    auth_header = base64.b64encode(credentials.encode()).decode()

    method = f'''
        /// <summary>
        /// {endpoint.description}
        /// </summary>
        /// <returns>{endpoint.return_description}</returns>
        [ExcelFunction(
            Name = "{endpoint.name}",
            Description = "{endpoint.description}",
            Category = "Domino Model APIs",
            IsVolatile = false,
            IsExceptionSafe = true
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


def generate_csharp_code() -> str:
    """Generate the complete C# add-in code."""

    methods = "\n".join([generate_udf_method(ep) for ep in ENDPOINTS])

    code = f'''using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

/// <summary>
/// Excel-DNA Add-in providing UDFs for Domino Model API endpoints.
///
/// This add-in was auto-generated and provides the following functions:
/// {chr(10).join(["/// - " + ep.name + ": " + ep.description[:60] + "..." for ep in ENDPOINTS])}
///
/// Each function calls a specific Domino model endpoint with the provided parameters
/// and returns the result from the model (supports array spilling for multiple results).
/// </summary>
public static class DominoModelFunctions
{{
    /// <summary>
    /// Extracts the "result" field from the JSON response and returns it as Excel-friendly output.
    /// Handles single values, 1D arrays (horizontal spill), and 2D arrays (grid spill).
    /// </summary>
    private static object ParseResult(string json)
    {{
        // Find the "result" field using regex (lightweight, no external JSON dependency)
        var resultMatch = Regex.Match(json, @"""result""\s*:\s*(\[[\s\S]*?\]|[^,\}}]+)");
        if (!resultMatch.Success)
        {{
            return "Error: No result field in response";
        }}

        string resultValue = resultMatch.Groups[1].Value.Trim();

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


def generate_dna_file() -> str:
    """Generate the Excel-DNA .dna configuration file."""
    return '''<?xml version="1.0" encoding="utf-8"?>
<DnaLibrary Name="Domino Model APIs Add-In" RuntimeVersion="v4.0">
  <ExternalLibrary Path="DominoModelFunctions.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" />
</DnaLibrary>
'''


def build_addin():
    """Build the Excel-DNA add-in."""

    print("=" * 60)
    print("Domino Model APIs Excel Add-in Generator")
    print("=" * 60)
    print()

    # Create a temporary build directory
    build_dir = tempfile.mkdtemp(prefix="exceldna_build_")
    print(f"[1/6] Created temporary build directory: {build_dir}")

    try:
        # Write the C# code
        cs_file = os.path.join(build_dir, "DominoModelFunctions.cs")
        with open(cs_file, "w") as f:
            f.write(generate_csharp_code())
        print(f"[2/6] Generated C# source code with {len(ENDPOINTS)} UDF(s):")
        for ep in ENDPOINTS:
            params = ", ".join([p["name"] for p in ep.parameters])
            print(f"       - {ep.name}({params})")

        # Write the .dna file
        dna_file = os.path.join(build_dir, "DominoModelFunctions.dna")
        with open(dna_file, "w") as f:
            f.write(generate_dna_file())
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

        # Find the PACKED .xll files - these are in publish folder with -packed suffix
        # The packed version has the DLL embedded and is what Excel actually needs
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
        for ep in ENDPOINTS:
            params = ", ".join([p["name"] for p in ep.parameters])
            print(f"  ={ep.name}({params})")
            print(f"    {ep.description[:70]}...")
            print()

        return copied_files[0][1] if copied_files else None

    finally:
        # Clean up the temp directory
        shutil.rmtree(build_dir, ignore_errors=True)


if __name__ == "__main__":
    try:
        build_addin()
    except Exception as e:
        print(f"\nError: {e}")
        print("\nTroubleshooting:")
        print("  - Ensure .NET SDK (6.0+) is installed: dotnet --version")
        print("  - On Linux/Mac, you may need to build on Windows for .xll generation")
        print("  - Alternatively, copy the generated .cs file to a Windows machine with Visual Studio")

        # Output the C# code so user can build manually if needed
        print("\n" + "=" * 60)
        print("FALLBACK: Generated C# Code (copy to build manually)")
        print("=" * 60)
        print(generate_csharp_code())
        print("\n" + "=" * 60)
        print("FALLBACK: .dna Configuration File")
        print("=" * 60)
        print(generate_dna_file())
