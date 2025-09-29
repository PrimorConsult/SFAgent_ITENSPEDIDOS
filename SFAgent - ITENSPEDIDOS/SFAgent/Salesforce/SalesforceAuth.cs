using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Collections.Generic;
using SFAgent.Config;

namespace SFAgent.Salesforce
{
    public class SalesforceTokenResponse
    {
        public string access_token { get; set; }
        public string instance_url { get; set; }
        public string token_type { get; set; }
        public int expires_in { get; set; }
    }

    public class SalesforceAuth
    {
        private string _token;
        private DateTime _tokenExpiration;

        public async Task<string> GetValidToken()
        {
            if (!string.IsNullOrEmpty(_token) && DateTime.Now < _tokenExpiration)
                return _token;

            using (var client = new HttpClient())
            {
                var body = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("grant_type", "client_credentials"),
                    new KeyValuePair<string, string>("client_id", ConfigCred.ClientId),
                    new KeyValuePair<string, string>("client_secret", ConfigCred.ClientSecret)
                });

                var response = await client.PostAsync(ConfigUrls.AuthUrl, body);
                var json = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                    throw new Exception($"Erro ao obter token: {json}");

                var tokenResponse = JsonConvert.DeserializeObject<SalesforceTokenResponse>(json);
                _token = tokenResponse.access_token;
                _tokenExpiration = DateTime.Now.AddSeconds(tokenResponse.expires_in - 60);

                return _token;
            }
        }
    }
}