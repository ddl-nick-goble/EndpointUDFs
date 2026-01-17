using System;
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
/// - Domino.GetCreditCurvesApi: Calls the GetCreditCurvesApi Domino model API endpoint....
/// - Domino.GetLoanProbabilityOfDefaultApi: Calls the GetLoanProbabilityOfDefaultApi Domino model API en...
/// - Domino.GetExpectedLossApi: Calls the GetExpectedLossApi Domino model API endpoint....
/// - Domino.GetLoanInventoryApi: Calls the GetLoanInventoryApi Domino model API endpoint....
///
/// Each function calls a specific Domino model endpoint with the provided parameters
/// and returns the result from the model (supports array spilling for multiple results).
/// </summary>
public static class DominoModelFunctions
{
    private enum ParamKind
    {
        String,
        Number,
        Bool,
        Date,
    }

    /// <summary>
    /// Formats a date-like parameter into yyyy-MM-dd.
    /// Accepts Excel dates, Unix epoch (seconds/ms), and date strings.
    /// </summary>
    private static string FormatDateParam(object value)
    {
        if (value == null)
        {
            return "";
        }
        if (value is ExcelMissing || value is ExcelEmpty)
        {
            return "";
        }

        if (value is double d)
        {
            return FormatDateFromNumber(d);
        }
        if (value is int i)
        {
            return FormatDateFromNumber(i);
        }
        if (value is DateTime dt)
        {
            return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }
        if (value is string s)
        {
            s = s.Trim();
            if (string.IsNullOrEmpty(s))
            {
                return "";
            }

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
            {
                return FormatDateFromNumber(num);
            }

            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime parsed))
            {
                return parsed.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            return s;
        }

        return value.ToString();
    }

    private static string FormatDateFromNumber(double value)
    {
        // Epoch milliseconds or seconds
        if (value >= 1_000_000_000_000d)
        {
            var dt = DateTimeOffset.FromUnixTimeMilliseconds((long)Math.Round(value)).DateTime;
            return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }
        if (value >= 1_000_000_000d)
        {
            var dt = DateTimeOffset.FromUnixTimeSeconds((long)Math.Round(value)).DateTime;
            return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        try
        {
            var dt = DateTime.FromOADate(value);
            return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }
        catch
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }
    }

    private static string EscapeJsonString(string value)
    {
        if (value == null)
        {
            return "";
        }
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    private static object NormalizeExcelValue(object value)
    {
        if (value is ExcelReference excelRef)
        {
            try
            {
                value = excelRef.GetValue();
            }
            catch
            {
                return value;
            }
        }
        return value;
    }

    private static string SerializeParamValue(object value, ParamKind kind, bool allowArray)
    {
        value = NormalizeExcelValue(value);
        if (value == null || value is ExcelMissing || value is ExcelEmpty)
        {
            return "null";
        }

        if (allowArray && value is Array array)
        {
            return SerializeArrayValue(array, kind);
        }

        if (!allowArray && value is Array singleCell && singleCell.Rank == 2 && singleCell.GetLength(0) == 1 && singleCell.GetLength(1) == 1)
        {
            return SerializeScalarValue(singleCell.GetValue(0, 0), kind);
        }

        return SerializeScalarValue(value, kind);
    }

    private static bool IsEmptyCell(object value)
    {
        return value == null || value is ExcelMissing || value is ExcelEmpty || value is ExcelError;
    }

    private static string SerializeArrayValue(Array array, ParamKind kind)
    {
        if (array.Rank == 1)
        {
            var sb = new StringBuilder();
            sb.Append('[');
            int start = array.GetLowerBound(0);
            int end = array.GetUpperBound(0);
            bool appended = false;
            for (int i = start; i <= end; i++)
            {
                object value = array.GetValue(i);
                if (IsEmptyCell(value))
                {
                    continue;
                }
                if (appended)
                {
                    sb.Append(',');
                }
                sb.Append(kind == ParamKind.Number ? SerializeScalarValue(Convert.ToString(value, CultureInfo.InvariantCulture), ParamKind.String) : SerializeScalarValue(value, kind));
                appended = true;
            }
            sb.Append(']');
            return sb.ToString();
        }

        if (array.Rank == 2)
        {
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);
            var sb = new StringBuilder();
            sb.Append('[');

            if (rows == 1 || cols == 1)
            {
                int count = rows == 1 ? cols : rows;
                bool appended = false;
                for (int i = 0; i < count; i++)
                {
                    object value = rows == 1 ? array.GetValue(0, i) : array.GetValue(i, 0);
                    if (IsEmptyCell(value))
                    {
                        continue;
                    }
                    if (appended)
                    {
                        sb.Append(',');
                    }
                    sb.Append(kind == ParamKind.Number ? SerializeScalarValue(Convert.ToString(value, CultureInfo.InvariantCulture), ParamKind.String) : SerializeScalarValue(value, kind));
                    appended = true;
                }
                sb.Append(']');
                return sb.ToString();
            }

            for (int r = 0; r < rows; r++)
            {
                if (r > 0)
                {
                    sb.Append(',');
                }
                sb.Append('[');
                bool appended = false;
                for (int c = 0; c < cols; c++)
                {
                    object value = array.GetValue(r, c);
                    if (IsEmptyCell(value))
                    {
                        continue;
                    }
                    if (appended)
                    {
                        sb.Append(',');
                    }
                    sb.Append(kind == ParamKind.Number ? SerializeScalarValue(Convert.ToString(value, CultureInfo.InvariantCulture), ParamKind.String) : SerializeScalarValue(value, kind));
                    appended = true;
                }
                sb.Append(']');
            }
            sb.Append(']');
            return sb.ToString();
        }

        return SerializeScalarValue(array, kind);
    }

    private static string SerializeScalarValue(object value, ParamKind kind)
    {
        if (value == null || value is ExcelMissing || value is ExcelEmpty)
        {
            return "null";
        }

        switch (kind)
        {
            case ParamKind.String:
                return "\"" + EscapeJsonString(Convert.ToString(value, CultureInfo.InvariantCulture)) + "\"";
            case ParamKind.Date:
                return "\"" + EscapeJsonString(FormatDateParam(value)) + "\"";
            case ParamKind.Bool:
                if (value is bool b)
                {
                    return b ? "true" : "false";
                }
                if (value is double d)
                {
                    return d != 0d ? "true" : "false";
                }
                if (value is int i)
                {
                    return i != 0 ? "true" : "false";
                }
                if (value is string s)
                {
                    if (bool.TryParse(s, out bool parsedBool))
                    {
                        return parsedBool ? "true" : "false";
                    }
                    if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedNum))
                    {
                        return parsedNum != 0d ? "true" : "false";
                    }
                }
                return "false";
            default:
                return SerializeNumberValue(value, false);
        }
    }

    private static string SerializeNumberValue(object value, bool forceFloat)
    {
        if (value is double num)
        {
            return FormatNumber(num, forceFloat);
        }
        if (value is int numInt)
        {
            return FormatNumber(numInt, forceFloat);
        }
        if (value is string str)
        {
            if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed))
            {
                return FormatNumber(parsed, forceFloat);
            }
            return "\"" + EscapeJsonString(str) + "\"";
        }
        try
        {
            return FormatNumber(Convert.ToDouble(value, CultureInfo.InvariantCulture), forceFloat);
        }
        catch
        {
            return "\"" + EscapeJsonString(Convert.ToString(value, CultureInfo.InvariantCulture)) + "\"";
        }
    }

    private static string FormatNumber(double value, bool forceFloat)
    {
        if (!forceFloat)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }
        if (Math.Abs(value % 1) < 1e-12)
        {
            return value.ToString("0.0", CultureInfo.InvariantCulture);
        }
        return value.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Extracts the "result" field from the JSON response and returns it as Excel-friendly output.
    /// Handles single values, 1D arrays (horizontal spill), and 2D arrays (grid spill).
    /// </summary>
    private static object ParseResult(string json)
    {
        if (!TryExtractResultValue(json, out string resultValue, out string error))
        {
            return error;
        }

        // Check if it's a 2D array (array of arrays) like [[1,2],[3,4]]
        if (resultValue.StartsWith("[["))
        {
            return Parse2DArray(resultValue);
        }
        // Check if it's a 1D array like [1,2,3]
        else if (resultValue.StartsWith("[") && resultValue.EndsWith("]"))
        {
            return Parse1DArray(resultValue);
        }
        else
        {
            // Single value
            return ParseSingleValue(resultValue);
        }
    }

    /// <summary>
    /// Extracts the raw JSON value for the "result" field without relying on regex for nested arrays.
    /// </summary>
    private static bool TryExtractResultValue(string json, out string resultValue, out string error)
    {
        resultValue = "";
        error = "";

        var match = Regex.Match(json, @"""result""\s*:");
        if (!match.Success)
        {
            error = "Error: No result field in response";
            return false;
        }

        int i = match.Index + match.Length;
        while (i < json.Length && char.IsWhiteSpace(json[i]))
        {
            i++;
        }

        if (i >= json.Length)
        {
            error = "Error: Empty result field in response";
            return false;
        }

        char start = json[i];
        if (start == '[' || start == '{')
        {
            char open = start;
            char close = (start == '[') ? ']' : '}';
            int depth = 0;
            bool inString = false;
            bool escape = false;
            int startIndex = i;

            for (; i < json.Length; i++)
            {
                char ch = json[i];
                if (inString)
                {
                    if (escape)
                    {
                        escape = false;
                        continue;
                    }
                    if (ch == '\\')
                    {
                        escape = true;
                        continue;
                    }
                    if (ch == '"')
                    {
                        inString = false;
                    }
                    continue;
                }

                if (ch == '"')
                {
                    inString = true;
                    continue;
                }

                if (ch == open)
                {
                    depth++;
                }
                else if (ch == close)
                {
                    depth--;
                    if (depth == 0)
                    {
                        resultValue = json.Substring(startIndex, i - startIndex + 1).Trim();
                        return true;
                    }
                }
            }

            error = "Error: Unterminated result value in response";
            return false;
        }

        if (start == '"')
        {
            int startIndex = i;
            bool escape = false;
            for (i = i + 1; i < json.Length; i++)
            {
                char ch = json[i];
                if (escape)
                {
                    escape = false;
                    continue;
                }
                if (ch == '\\')
                {
                    escape = true;
                    continue;
                }
                if (ch == '"')
                {
                    resultValue = json.Substring(startIndex, i - startIndex + 1).Trim();
                    return true;
                }
            }
            error = "Error: Unterminated string result in response";
            return false;
        }

        int primitiveStart = i;
        while (i < json.Length && json[i] != ',' && json[i] != '}' && json[i] != ']')
        {
            i++;
        }
        resultValue = json.Substring(primitiveStart, i - primitiveStart).Trim();
        return true;
    }

    /// <summary>
    /// Parses a 2D array like [[1,2,3],[4,5,6]] into an Excel-compatible object[,] for grid spill.
    /// </summary>
    private static object Parse2DArray(string arrayStr)
    {
        // Extract inner arrays using regex to find each [...] row
        var rowMatches = Regex.Matches(arrayStr, @"\[([^\[\]]*)\]");
        if (rowMatches.Count == 0)
        {
            return "Error: Invalid 2D array format";
        }

        // Parse each row to get dimensions and values
        var rows = new List<List<object>>();
        int maxCols = 0;

        foreach (Match rowMatch in rowMatches)
        {
            string rowContent = rowMatch.Groups[1].Value;
            var rowValues = new List<object>();

            if (!string.IsNullOrWhiteSpace(rowContent))
            {
                var parts = rowContent.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                {
                    rowValues.Add(ParseSingleValue(part.Trim()));
                }
            }

            rows.Add(rowValues);
            if (rowValues.Count > maxCols)
            {
                maxCols = rowValues.Count;
            }
        }

        // Handle edge cases
        if (rows.Count == 0 || maxCols == 0)
        {
            return "";
        }
        if (rows.Count == 1 && rows[0].Count == 1)
        {
            return rows[0][0];
        }

        // Create 2D array for Excel grid spill
        object[,] spillArray = new object[rows.Count, maxCols];
        for (int r = 0; r < rows.Count; r++)
        {
            for (int c = 0; c < maxCols; c++)
            {
                if (c < rows[r].Count)
                {
                    spillArray[r, c] = rows[r][c];
                }
                else
                {
                    spillArray[r, c] = ""; // Pad jagged arrays
                }
            }
        }
        return spillArray;
    }

    /// <summary>
    /// Parses a 1D array like [1,2,3] into an Excel-compatible object[,] for horizontal spill.
    /// </summary>
    private static object Parse1DArray(string arrayStr)
    {
        string inner = arrayStr.Substring(1, arrayStr.Length - 2).Trim();
        if (string.IsNullOrEmpty(inner))
        {
            return "";
        }

        var parts = inner.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        var results = new List<object>();

        foreach (var part in parts)
        {
            results.Add(ParseSingleValue(part.Trim()));
        }

        if (results.Count == 1)
        {
            return results[0];
        }

        // Create a 1-row, N-column array for horizontal spill
        object[,] spillArray = new object[1, results.Count];
        for (int i = 0; i < results.Count; i++)
        {
            spillArray[0, i] = results[i];
        }
        return spillArray;
    }

    /// <summary>
    /// Parses a single value (number or string) into the appropriate type.
    /// </summary>
    private static object ParseSingleValue(string value)
    {
        string trimmed = value.Trim();
        if (double.TryParse(trimmed, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out double numVal))
        {
            return numVal;
        }
        return trimmed.Trim('"');
    }


        /// <summary>
        /// Calls the GetCreditCurvesApi Domino model API endpoint.
        /// </summary>
        /// <returns>Returns the model result value (spills across cells if array)</returns>
        [ExcelFunction(
            Name = "Domino.GetCreditCurvesApi",
            Description = "Calls the GetCreditCurvesApi Domino model API endpoint.",
            Category = "Domino Model APIs",
            IsVolatile = false,
            IsExceptionSafe = true,
            IsThreadSafe = true,
            HelpTopic = "https://se-demo.domino.tech/models/696a8cfede2f14747436d212/overview?ownerName=nick_goble&projectName=EndpointUDFs"
        )]
        public static object GetCreditCurvesApi(
            [ExcelArgument(Name = "curve_date", Description = "The curve_date parameter for the model", AllowReference = true)] object curve_date)
        {
            try
            {
                // Force TLS 1.2 (required for modern HTTPS endpoints)
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

                string url = "https://se-demo.domino.tech:443/models/696a8cfede2f14747436d212/latest/model";
                string jsonPayload = "{\"data\": {" + "\"curve_date\": " + SerializeParamValue(curve_date, ParamKind.Date, false) + "" + "}}";

                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.ContentType] = "application/json";
                    client.Headers[HttpRequestHeader.Authorization] = "Basic bkRJamMxekFWSE1iSmU0ZTdQalYxbGxVMVpSc3p6OXI0S1l5R3J1UU1Bb3BUVTJZcDU5VmdURmlJZGc5aHVxSDpuRElqYzF6QVZITWJKZTRlN1BqVjFsbFUxWlJzeno5cjRLWXlHcnVRTUFvcFRVMllwNTlWZ1RGaUlkZzlodXFI";

                    string response = client.UploadString(url, "POST", jsonPayload);
                    return ParseResult(response);
                }
            }
            catch (WebException ex)
            {
                if (ex.Response != null)
                {
                    using (var reader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        return "API Error: " + reader.ReadToEnd();
                    }
                }
                return "Error: " + ex.Message;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }


        /// <summary>
        /// Calls the GetLoanProbabilityOfDefaultApi Domino model API endpoint.
        /// </summary>
        /// <returns>Returns the model result value (spills across cells if array)</returns>
        [ExcelFunction(
            Name = "Domino.GetLoanProbabilityOfDefaultApi",
            Description = "Calls the GetLoanProbabilityOfDefaultApi Domino model API endpoint.",
            Category = "Domino Model APIs",
            IsVolatile = false,
            IsExceptionSafe = true,
            IsThreadSafe = true,
            HelpTopic = "https://se-demo.domino.tech/models/696a8cffde2f14747436d217/overview?ownerName=nick_goble&projectName=EndpointUDFs"
        )]
        public static object GetLoanProbabilityOfDefaultApi(
            [ExcelArgument(Name = "loan_id", Description = "The loan_id parameter for the model", AllowReference = false)] object loan_id, [ExcelArgument(Name = "credit_score", Description = "The credit_score parameter for the model", AllowReference = false)] object credit_score, [ExcelArgument(Name = "debt_to_income_ratio", Description = "The debt_to_income_ratio parameter for the model", AllowReference = false)] object debt_to_income_ratio, [ExcelArgument(Name = "loan_to_value_ratio", Description = "The loan_to_value_ratio parameter for the model", AllowReference = false)] object loan_to_value_ratio, [ExcelArgument(Name = "loan_age_months", Description = "The loan_age_months parameter for the model", AllowReference = false)] object loan_age_months, [ExcelArgument(Name = "original_principal_balance", Description = "The original_principal_balance parameter for the model", AllowReference = false)] object original_principal_balance, [ExcelArgument(Name = "interest_rate", Description = "The interest_rate parameter for the model", AllowReference = false)] object interest_rate, [ExcelArgument(Name = "employment_years", Description = "The employment_years parameter for the model", AllowReference = false)] object employment_years, [ExcelArgument(Name = "delinquency_30d_past_12m", Description = "The delinquency_30d_past_12m parameter for the model", AllowReference = false)] object delinquency_30d_past_12m, [ExcelArgument(Name = "loan_purpose", Description = "The loan_purpose parameter for the model", AllowReference = false)] object loan_purpose)
        {
            try
            {
                // Force TLS 1.2 (required for modern HTTPS endpoints)
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

                string url = "https://se-demo.domino.tech:443/models/696a8cffde2f14747436d217/latest/model";
                string jsonPayload = "{\"data\": {" + "\"loan_id\": " + SerializeParamValue(loan_id, ParamKind.String, false) + ", " + "\"credit_score\": " + SerializeParamValue(credit_score, ParamKind.String, false) + ", " + "\"debt_to_income_ratio\": " + SerializeParamValue(debt_to_income_ratio, ParamKind.String, false) + ", " + "\"loan_to_value_ratio\": " + SerializeParamValue(loan_to_value_ratio, ParamKind.String, false) + ", " + "\"loan_age_months\": " + SerializeParamValue(loan_age_months, ParamKind.String, false) + ", " + "\"original_principal_balance\": " + SerializeParamValue(original_principal_balance, ParamKind.String, false) + ", " + "\"interest_rate\": " + SerializeParamValue(interest_rate, ParamKind.String, false) + ", " + "\"employment_years\": " + SerializeParamValue(employment_years, ParamKind.String, false) + ", " + "\"delinquency_30d_past_12m\": " + SerializeParamValue(delinquency_30d_past_12m, ParamKind.String, false) + ", " + "\"loan_purpose\": " + SerializeParamValue(loan_purpose, ParamKind.String, false) + "" + "}}";

                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.ContentType] = "application/json";
                    client.Headers[HttpRequestHeader.Authorization] = "Basic MVdCY2pQWkRQcjNhUTlCZUM4ZHpTNTVIUlo1MWcwUkppc3ZpaWgwNURBZkVPSFVGTld6VERqb2J2d3dhQW9mNDoxV0JjalBaRFByM2FROUJlQzhkelM1NUhSWjUxZzBSSmlzdmlpaDA1REFmRU9IVUZOV3pURGpvYnZ3d2FBb2Y0";

                    string response = client.UploadString(url, "POST", jsonPayload);
                    return ParseResult(response);
                }
            }
            catch (WebException ex)
            {
                if (ex.Response != null)
                {
                    using (var reader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        return "API Error: " + reader.ReadToEnd();
                    }
                }
                return "Error: " + ex.Message;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }


        /// <summary>
        /// Calls the GetExpectedLossApi Domino model API endpoint.
        /// </summary>
        /// <returns>Returns the model result value (spills across cells if array)</returns>
        [ExcelFunction(
            Name = "Domino.GetExpectedLossApi",
            Description = "Calls the GetExpectedLossApi Domino model API endpoint.",
            Category = "Domino Model APIs",
            IsVolatile = false,
            IsExceptionSafe = true,
            IsThreadSafe = true,
            HelpTopic = "https://se-demo.domino.tech/models/696a8cffde2f14747436d21c/overview?ownerName=nick_goble&projectName=EndpointUDFs"
        )]
        public static object GetExpectedLossApi(
            [ExcelArgument(Name = "loan_id", Description = "The loan_id parameter for the model", AllowReference = false)] object loan_id, [ExcelArgument(Name = "probability_of_default_1y", Description = "The probability_of_default_1y parameter for the model", AllowReference = false)] object probability_of_default_1y, [ExcelArgument(Name = "exposure_at_default", Description = "The exposure_at_default parameter for the model", AllowReference = false)] object exposure_at_default, [ExcelArgument(Name = "loss_given_default", Description = "The loss_given_default parameter for the model", AllowReference = false)] object loss_given_default, [ExcelArgument(Name = "credit_rating", Description = "The credit_rating parameter for the model", AllowReference = false)] object credit_rating, [ExcelArgument(Name = "remaining_term_years", Description = "The remaining_term_years parameter for the model", AllowReference = false)] object remaining_term_years, [ExcelArgument(Name = "curve_tenors", Description = "The curve_tenors parameter for the model", AllowReference = true)] object curve_tenors, [ExcelArgument(Name = "curve_rates", Description = "The curve_rates parameter for the model", AllowReference = true)] object curve_rates)
        {
            try
            {
                // Force TLS 1.2 (required for modern HTTPS endpoints)
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

                string url = "https://se-demo.domino.tech:443/models/696a8cffde2f14747436d21c/latest/model";
                string jsonPayload = "{\"data\": {" + "\"loan_id\": " + SerializeParamValue(loan_id, ParamKind.String, false) + ", " + "\"probability_of_default_1y\": " + SerializeParamValue(probability_of_default_1y, ParamKind.String, false) + ", " + "\"exposure_at_default\": " + SerializeParamValue(exposure_at_default, ParamKind.String, false) + ", " + "\"loss_given_default\": " + SerializeParamValue(loss_given_default, ParamKind.String, false) + ", " + "\"credit_rating\": " + SerializeParamValue(credit_rating, ParamKind.String, false) + ", " + "\"remaining_term_years\": " + SerializeParamValue(remaining_term_years, ParamKind.String, false) + ", " + "\"curve_tenors\": " + SerializeParamValue(curve_tenors, ParamKind.Number, true) + ", " + "\"curve_rates\": " + SerializeParamValue(curve_rates, ParamKind.Number, true) + "" + "}}";

                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.ContentType] = "application/json";
                    client.Headers[HttpRequestHeader.Authorization] = "Basic UEozQm80WjZJeXZiUGp0ZE5qMGo5cDRVdTRsRzh5bVhVSlo5U2R3RmdoeVNHeW4wc20wV2tZeXpqRlkyVjRlUzpQSjNCbzRaNkl5dmJQanRkTmowajlwNFV1NGxHOHltWFVKWjlTZHdGZ2h5U0d5bjBzbTBXa1l5empGWTJWNGVT";

                    string response = client.UploadString(url, "POST", jsonPayload);
                    return ParseResult(response);
                }
            }
            catch (WebException ex)
            {
                if (ex.Response != null)
                {
                    using (var reader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        return "API Error: " + reader.ReadToEnd();
                    }
                }
                return "Error: " + ex.Message;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }


        /// <summary>
        /// Calls the GetLoanInventoryApi Domino model API endpoint.
        /// </summary>
        /// <returns>Returns the model result value (spills across cells if array)</returns>
        [ExcelFunction(
            Name = "Domino.GetLoanInventoryApi",
            Description = "Calls the GetLoanInventoryApi Domino model API endpoint.",
            Category = "Domino Model APIs",
            IsVolatile = false,
            IsExceptionSafe = true,
            IsThreadSafe = true,
            HelpTopic = "https://se-demo.domino.tech/models/696bfe46de2f14747436d2a2/overview?ownerName=nick_goble&projectName=EndpointUDFs"
        )]
        public static object GetLoanInventoryApi(
            [ExcelArgument(Name = "inventory_date", Description = "The inventory_date parameter for the model", AllowReference = true)] object inventory_date)
        {
            try
            {
                // Force TLS 1.2 (required for modern HTTPS endpoints)
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

                string url = "https://se-demo.domino.tech:443/models/696bfe46de2f14747436d2a2/latest/model";
                string jsonPayload = "{\"data\": {" + "\"inventory_date\": " + SerializeParamValue(inventory_date, ParamKind.Date, false) + "" + "}}";

                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.ContentType] = "application/json";
                    client.Headers[HttpRequestHeader.Authorization] = "Basic WDRnRUt2VGE0QUdlOHBxTmxEZGI5aW5EUnk0VkdXREtSeEMxNk5rWjdlUEN0WktuUWowZXpjUTZIVzBwbGJlVDpYNGdFS3ZUYTRBR2U4cHFObERkYjlpbkRSeTRWR1dES1J4QzE2TmtaN2VQQ3RaS25RajBlemNRNkhXMHBsYmVU";

                    string response = client.UploadString(url, "POST", jsonPayload);
                    return ParseResult(response);
                }
            }
            catch (WebException ex)
            {
                if (ex.Response != null)
                {
                    using (var reader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        return "API Error: " + reader.ReadToEnd();
                    }
                }
                return "Error: " + ex.Message;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

}
