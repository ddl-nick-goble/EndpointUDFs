using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ExcelDna.Integration;

namespace DominoUdfs
{
    public static class DominoUdfs
    {
        private static readonly HttpClient Client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(30)
        };

        private static string CallEndpoint(string url, string user, string pass, object payload)
        {
            try
            {
                var json = JsonSerializer.Serialize(payload);
                using var request = new HttpRequestMessage(HttpMethod.Post, url);
                var token = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{user}:{pass}"));
                request.Headers.Authorization = new AuthenticationHeaderValue("Basic", token);
                request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                using var response = Client.SendAsync(request).GetAwaiter().GetResult();
                var body = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                return $"{(int)response.StatusCode} {response.StatusCode}: {body}";
            }
            catch (Exception ex)
            {
                return $"ERROR: {ex.Message}";
            }
        }

        [ExcelFunction(
            Name = "FirstTestEndpoint",
            Description = "Calls Domino endpoint FirstTestEndpoint and returns status + response body."
        )]
        public static string FirstTestEndpoint()
        {
            var payload = new Dictionary<string, object> {
                {"data", new Dictionary<string, object> {  } }
            };
            return CallEndpoint("", "", "", payload);
        }

        [ExcelFunction(
            Name = "SecondTestEndpoint",
            Description = "Calls Domino endpoint SecondTestEndpoint and returns status + response body."
        )]
        public static string SecondTestEndpoint()
        {
            var payload = new Dictionary<string, object> {
                {"data", new Dictionary<string, object> {  } }
            };
            return CallEndpoint("", "", "", payload);
        }

    }
}