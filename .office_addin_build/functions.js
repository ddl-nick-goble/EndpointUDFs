/**
 * Domino Model APIs - Office Add-in Custom Functions
 * Auto-generated JavaScript implementations for calling Domino ML model endpoints
 */

// Helper function to format date parameters
function formatDateParam(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  // If it's already a string in yyyy-MM-dd format, return as-is
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value)) {
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

