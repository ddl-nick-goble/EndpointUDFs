"""
Domino Office Add-in Static File Server
Auto-generated Flask app to serve Office Add-in files
"""

import requests
from flask import Flask, Response, request

app = Flask(__name__)

def get_mimetype(filename):
    """Get MIME type based on file extension."""
    if filename.endswith('.xml'):
        return 'application/xml'
    elif filename.endswith('.json'):
        return 'application/json'
    elif filename.endswith('.js'):
        return 'application/javascript'
    elif filename.endswith('.html'):
        return 'text/html'
    else:
        return 'text/plain'

# Embedded files (generated at build time)
FILES = {
    'manifest.xml': ('''<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictaskpane/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">

  <!-- Unique identifier for this add-in -->
  <Id>d083a7e1-92e5-4b82-9c71-86eeae482fe6</Id>

  <!-- General information -->
  <Version>1.0.0.0</Version>
  <ProviderName>Domino Data Lab</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Domino Model APIs - EndpointUDFs"/>
  <Description DefaultValue="Custom functions for calling Domino ML model endpoints from Excel"/>
  <IconUrl DefaultValue="https://www.dominodatalab.com/favicon.ico"/>
  <HighResolutionIconUrl DefaultValue="https://www.dominodatalab.com/favicon.ico"/>
  <SupportUrl DefaultValue="https://www.dominodatalab.com"/>

  <!-- Specify which Office applications support this add-in -->
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>

  <!-- Default settings -->
  <DefaultSettings>
    <SourceLocation DefaultValue="https://se-demo.domino.tech:443/models/PLACEHOLDER/latest/model?file=index.html"/>
  </DefaultSettings>

  <!-- Permissions required -->
  <Permissions>ReadWriteDocument</Permissions>

  <!-- Version overrides for custom functions -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation DefaultValue="https://se-demo.domino.tech:443/models/PLACEHOLDER/latest/model?file=functions.js"/>
            </Script>
            <Page>
              <SourceLocation DefaultValue="https://se-demo.domino.tech:443/models/PLACEHOLDER/latest/model?file=index.html"/>
            </Page>
            <Metadata>
              <SourceLocation DefaultValue="https://se-demo.domino.tech:443/models/PLACEHOLDER/latest/model?file=functions.json"/>
            </Metadata>
            <Namespace resid="Functions.Namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://www.dominodatalab.com/favicon.ico"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://www.dominodatalab.com/favicon.ico"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://www.dominodatalab.com/favicon.ico"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Functions.Script.Url" DefaultValue="https://se-demo.domino.tech:443/models/PLACEHOLDER/latest/model?file=functions.js"/>
        <bt:Url id="Functions.Metadata.Url" DefaultValue="https://se-demo.domino.tech:443/models/PLACEHOLDER/latest/model?file=functions.json"/>
        <bt:Url id="Functions.Page.Url" DefaultValue="https://se-demo.domino.tech:443/models/PLACEHOLDER/latest/model?file=index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Functions.Namespace" DefaultValue="DOMINO"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Functions.Description" DefaultValue="Custom functions for Domino ML model endpoints"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
''', get_mimetype('manifest.xml')),
    'functions.json': ('''{
  "functions": [
    {
      "id": "GETCREDITCURVESAPI",
      "name": "DOMINO.ENDPOINTUDFS.GETCREDITCURVESAPI",
      "description": "Calls the GetCreditCurvesApi Domino model API endpoint.",
      "parameters": [
        {
          "name": "curve_date",
          "description": "The curve_date parameter for the model",
          "type": "any"
        }
      ],
      "result": {
        "type": "any",
        "dimensionality": "matrix"
      },
      "options": {
        "stream": false,
        "cancelable": false
      }
    },
    {
      "id": "GETLOANPROBABILITYOFDEFAULTAPI",
      "name": "DOMINO.ENDPOINTUDFS.GETLOANPROBABILITYOFDEFAULTAPI",
      "description": "Calls the GetLoanProbabilityOfDefaultApi Domino model API endpoint.",
      "parameters": [
        {
          "name": "loan_id",
          "description": "The loan_id parameter for the model",
          "type": "any"
        },
        {
          "name": "credit_score",
          "description": "The credit_score parameter for the model",
          "type": "any"
        },
        {
          "name": "debt_to_income_ratio",
          "description": "The debt_to_income_ratio parameter for the model",
          "type": "any"
        },
        {
          "name": "loan_to_value_ratio",
          "description": "The loan_to_value_ratio parameter for the model",
          "type": "any"
        },
        {
          "name": "loan_age_months",
          "description": "The loan_age_months parameter for the model",
          "type": "any"
        },
        {
          "name": "original_principal_balance",
          "description": "The original_principal_balance parameter for the model",
          "type": "any"
        },
        {
          "name": "interest_rate",
          "description": "The interest_rate parameter for the model",
          "type": "any"
        },
        {
          "name": "employment_years",
          "description": "The employment_years parameter for the model",
          "type": "any"
        },
        {
          "name": "delinquency_30d_past_12m",
          "description": "The delinquency_30d_past_12m parameter for the model",
          "type": "any"
        },
        {
          "name": "loan_purpose",
          "description": "The loan_purpose parameter for the model",
          "type": "any"
        }
      ],
      "result": {
        "type": "any",
        "dimensionality": "matrix"
      },
      "options": {
        "stream": false,
        "cancelable": false
      }
    },
    {
      "id": "GETEXPECTEDLOSSAPI",
      "name": "DOMINO.ENDPOINTUDFS.GETEXPECTEDLOSSAPI",
      "description": "Calls the GetExpectedLossApi Domino model API endpoint.",
      "parameters": [
        {
          "name": "loan_id",
          "description": "The loan_id parameter for the model",
          "type": "any"
        },
        {
          "name": "probability_of_default_1y",
          "description": "The probability_of_default_1y parameter for the model",
          "type": "any"
        },
        {
          "name": "exposure_at_default",
          "description": "The exposure_at_default parameter for the model",
          "type": "any"
        },
        {
          "name": "loss_given_default",
          "description": "The loss_given_default parameter for the model",
          "type": "any"
        },
        {
          "name": "credit_rating",
          "description": "The credit_rating parameter for the model",
          "type": "any"
        },
        {
          "name": "remaining_term_years",
          "description": "The remaining_term_years parameter for the model",
          "type": "any"
        },
        {
          "name": "curve_tenors",
          "description": "The curve_tenors parameter for the model",
          "type": "any",
          "dimensionality": "matrix"
        },
        {
          "name": "curve_rates",
          "description": "The curve_rates parameter for the model",
          "type": "any",
          "dimensionality": "matrix"
        }
      ],
      "result": {
        "type": "any",
        "dimensionality": "matrix"
      },
      "options": {
        "stream": false,
        "cancelable": false
      }
    },
    {
      "id": "GETLOANINVENTORYAPI",
      "name": "DOMINO.ENDPOINTUDFS.GETLOANINVENTORYAPI",
      "description": "Calls the GetLoanInventoryApi Domino model API endpoint.",
      "parameters": [
        {
          "name": "inventory_date",
          "description": "The inventory_date parameter for the model",
          "type": "any"
        }
      ],
      "result": {
        "type": "any",
        "dimensionality": "matrix"
      },
      "options": {
        "stream": false,
        "cancelable": false
      }
    }
  ]
}''', get_mimetype('functions.json')),
    'functions.js': ('''/**
 * Domino Model APIs - Office Add-in Custom Functions
 * Auto-generated JavaScript implementations for calling Domino ML model endpoints
 */

// Helper function to format date parameters
function formatDateParam(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  // If it's already a string in yyyy-MM-dd format, return as-is
  if (typeof value === 'string' && /^\\d{4}-\\d{2}-\\d{2}/.test(value)) {
    return value.substring(0, 10);
  }

  // If it's a Date object or Excel date number
  let date;
  if (value instanceof Date) {
    date = value;
  } else if (typeof value === 'number') {
    // Excel dates are days since 1900-01-01 (with a bug for 1900 being a leap year)
    // Convert to JavaScript Date
    const excelEpoch = new Date(1900, 0, 1);
    date = new Date(excelEpoch.getTime() + (value - 2) * 24 * 60 * 60 * 1000);
  } else {
    // Try to parse as string
    date = new Date(value);
  }

  // Format as yyyy-MM-dd
  if (date && !isNaN(date.getTime())) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  // Fallback to string representation
  return String(value);
}

// Function implementations

/**
 * Calls the GetCreditCurvesApi Domino model API endpoint.
 * @customfunction
 * @param {curve_date}
 * @returns {Promise<any>} The model prediction result
 */
async function getcreditcurvesapi(curve_date) {
  try {
    const targetUrl = "https://se-demo.domino.tech:443/models/696a8cfede2f14747436d212/latest/model";
    const payload = {
      data: {
        "curve_date": formatDateParam(curve_date)
      }
    };

    const response = await fetch("/proxy", {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ url: targetUrl, payload })
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`API Error: ${response.status} - ${errorText}`);
    }

    const result = await response.json();

    // Extract the result field
    if (result && 'result' in result) {
      return result.result;
    }

    return result;
  } catch (error) {
    return `#ERROR: ${error.message}`;
  }
}

// Register the function with CustomFunctions
CustomFunctions.associate("GETCREDITCURVESAPI", getcreditcurvesapi);

/**
 * Calls the GetLoanProbabilityOfDefaultApi Domino model API endpoint.
 * @customfunction
 * @param {loan_id, credit_score, debt_to_income_ratio, loan_to_value_ratio, loan_age_months, original_principal_balance, interest_rate, employment_years, delinquency_30d_past_12m, loan_purpose}
 * @returns {Promise<any>} The model prediction result
 */
async function getloanprobabilityofdefaultapi(loan_id, credit_score, debt_to_income_ratio, loan_to_value_ratio, loan_age_months, original_principal_balance, interest_rate, employment_years, delinquency_30d_past_12m, loan_purpose) {
  try {
    const targetUrl = "https://se-demo.domino.tech:443/models/696a8cffde2f14747436d217/latest/model";
    const payload = {
      data: {
        "loan_id": String(loan_id),
        "credit_score": String(credit_score),
        "debt_to_income_ratio": String(debt_to_income_ratio),
        "loan_to_value_ratio": String(loan_to_value_ratio),
        "loan_age_months": String(loan_age_months),
        "original_principal_balance": String(original_principal_balance),
        "interest_rate": String(interest_rate),
        "employment_years": String(employment_years),
        "delinquency_30d_past_12m": String(delinquency_30d_past_12m),
        "loan_purpose": String(loan_purpose)
      }
    };

    const response = await fetch("/proxy", {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ url: targetUrl, payload })
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`API Error: ${response.status} - ${errorText}`);
    }

    const result = await response.json();

    // Extract the result field
    if (result && 'result' in result) {
      return result.result;
    }

    return result;
  } catch (error) {
    return `#ERROR: ${error.message}`;
  }
}

// Register the function with CustomFunctions
CustomFunctions.associate("GETLOANPROBABILITYOFDEFAULTAPI", getloanprobabilityofdefaultapi);

/**
 * Calls the GetExpectedLossApi Domino model API endpoint.
 * @customfunction
 * @param {loan_id, probability_of_default_1y, exposure_at_default, loss_given_default, credit_rating, remaining_term_years, curve_tenors, curve_rates}
 * @returns {Promise<any>} The model prediction result
 */
async function getexpectedlossapi(loan_id, probability_of_default_1y, exposure_at_default, loss_given_default, credit_rating, remaining_term_years, curve_tenors, curve_rates) {
  try {
    const targetUrl = "https://se-demo.domino.tech:443/models/696a8cffde2f14747436d21c/latest/model";
    const payload = {
      data: {
        "loan_id": String(loan_id),
        "probability_of_default_1y": String(probability_of_default_1y),
        "exposure_at_default": String(exposure_at_default),
        "loss_given_default": String(loss_given_default),
        "credit_rating": String(credit_rating),
        "remaining_term_years": String(remaining_term_years),
        "curve_tenors": curve_tenors,
        "curve_rates": curve_rates
      }
    };

    const response = await fetch("/proxy", {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ url: targetUrl, payload })
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`API Error: ${response.status} - ${errorText}`);
    }

    const result = await response.json();

    // Extract the result field
    if (result && 'result' in result) {
      return result.result;
    }

    return result;
  } catch (error) {
    return `#ERROR: ${error.message}`;
  }
}

// Register the function with CustomFunctions
CustomFunctions.associate("GETEXPECTEDLOSSAPI", getexpectedlossapi);

/**
 * Calls the GetLoanInventoryApi Domino model API endpoint.
 * @customfunction
 * @param {inventory_date}
 * @returns {Promise<any>} The model prediction result
 */
async function getloaninventoryapi(inventory_date) {
  try {
    const targetUrl = "https://se-demo.domino.tech:443/models/696bfe46de2f14747436d2a2/latest/model";
    const payload = {
      data: {
        "inventory_date": formatDateParam(inventory_date)
      }
    };

    const response = await fetch("/proxy", {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ url: targetUrl, payload })
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`API Error: ${response.status} - ${errorText}`);
    }

    const result = await response.json();

    // Extract the result field
    if (result && 'result' in result) {
      return result.result;
    }

    return result;
  } catch (error) {
    return `#ERROR: ${error.message}`;
  }
}

// Register the function with CustomFunctions
CustomFunctions.associate("GETLOANINVENTORYAPI", getloaninventoryapi);

''', get_mimetype('functions.js')),
    'index.html': ('''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Domino Model APIs</title>

    <!-- Office.js -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            padding: 20px;
            line-height: 1.6;
        }
        h1 {
            color: #2b5797;
            font-size: 24px;
        }
        h2 {
            color: #4472C4;
            font-size: 18px;
            margin-top: 20px;
        }
        ul {
            list-style: none;
            padding: 0;
        }
        li {
            margin-bottom: 20px;
            padding: 10px;
            background: #f5f5f5;
            border-left: 4px solid #2b5797;
        }
        code {
            background: #e8e8e8;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: "Courier New", Courier, monospace;
        }
        .footer {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            color: #666;
            font-size: 12px;
        }
    </style>
</head>
<body>
    <h1>Domino Model APIs</h1>
    <p>Custom functions for calling Domino ML model endpoints from Excel.</p>

    <h2>Available Functions (4)</h2>
    <ul>

        <li>
          <strong>=<code>DOMINO.ENDPOINTUDFS.GETCREDITCURVESAPI(curve_date)</code></strong>
          <p>Calls the GetCreditCurvesApi Domino model API endpoint.</p>
        </li>

        <li>
          <strong>=<code>DOMINO.ENDPOINTUDFS.GETLOANPROBABILITYOFDEFAULTAPI(loan_id, credit_score, debt_to_income_ratio, loan_to_value_ratio, loan_age_months, original_principal_balance, interest_rate, employment_years, delinquency_30d_past_12m, loan_purpose)</code></strong>
          <p>Calls the GetLoanProbabilityOfDefaultApi Domino model API endpoint.</p>
        </li>

        <li>
          <strong>=<code>DOMINO.ENDPOINTUDFS.GETEXPECTEDLOSSAPI(loan_id, probability_of_default_1y, exposure_at_default, loss_given_default, credit_rating, remaining_term_years, curve_tenors, curve_rates)</code></strong>
          <p>Calls the GetExpectedLossApi Domino model API endpoint.</p>
        </li>

        <li>
          <strong>=<code>DOMINO.ENDPOINTUDFS.GETLOANINVENTORYAPI(inventory_date)</code></strong>
          <p>Calls the GetLoanInventoryApi Domino model API endpoint.</p>
        </li>
    </ul>

    <h2>Usage</h2>
    <p>Type any of the above functions into an Excel cell to call the corresponding Domino model endpoint.</p>
    <p><strong>Example:</strong> <code>=DOMINO.ENDPOINTUDFS.GETCREDITCURVESAPI(A1, B1)</code></p>

    <div class="footer">
        <p>Powered by Domino Data Lab | Auto-generated Office Add-in</p>
    </div>

    <script>
        Office.onReady(function() {
            console.log('Domino Model APIs add-in loaded');
        });
    </script>
</body>
</html>
''', get_mimetype('index.html'))
}

ALLOWED_URLS = {
    'https://se-demo.domino.tech:443/models/696a8cfede2f14747436d212/latest/model',
    'https://se-demo.domino.tech:443/models/696a8cffde2f14747436d217/latest/model',
    'https://se-demo.domino.tech:443/models/696a8cffde2f14747436d21c/latest/model',
    'https://se-demo.domino.tech:443/models/696bfe46de2f14747436d2a2/latest/model'
}

def _cors_response(response: Response) -> Response:
    origin = request.headers.get('Origin', '*')
    response.headers['Access-Control-Allow-Origin'] = origin
    response.headers['Access-Control-Allow-Credentials'] = 'true'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization, X-Domino-Api-Key'
    return response

@app.route('/', methods=['GET'])
def serve_file():
    """Serve files based on 'file' query parameter."""
    # Get the requested file from query parameter
    requested_file = request.args.get('file', 'index.html')

    if requested_file in FILES:
        content, mimetype = FILES[requested_file]
        response = Response(content, mimetype=mimetype)
        return _cors_response(response)

    return Response("File not found. Available files: " + ", ".join(FILES.keys()), status=404)

@app.route('/proxy', methods=['POST', 'OPTIONS'])
def proxy():
    """Proxy requests to Domino Model APIs using the user's session."""
    if request.method == 'OPTIONS':
        return _cors_response(Response(status=204))

    body = request.get_json(silent=True) or {}
    target_url = body.get('url')
    payload = body.get('payload')

    if not target_url or target_url not in ALLOWED_URLS:
        return _cors_response(Response("Invalid or missing target URL.", status=400))

    headers = {'Content-Type': 'application/json'}
    for header_name in ('Authorization', 'X-Domino-Api-Key'):
        header_value = request.headers.get(header_name)
        if header_value:
            headers[header_name] = header_value

    cookie_header = request.headers.get('Cookie')
    if cookie_header:
        headers['Cookie'] = cookie_header

    try:
        upstream = requests.post(target_url, headers=headers, json=payload, timeout=30)
    except requests.RequestException as exc:
        return _cors_response(Response(f"Upstream request failed: {exc}", status=502))

    response = Response(
        upstream.content,
        status=upstream.status_code,
        mimetype=upstream.headers.get('Content-Type', 'application/json'),
    )
    return _cors_response(response)

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint."""
    return {"status": "ok", "files": list(FILES.keys())}

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8888)
