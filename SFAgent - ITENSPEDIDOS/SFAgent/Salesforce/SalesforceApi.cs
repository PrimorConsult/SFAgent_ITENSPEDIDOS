using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SFAgent.Config;

namespace SFAgent.Salesforce
{
    public class SalesforceApi
    {
        private static readonly HttpClient _http = new HttpClient();

        // ---- Models ----
        public class UpsertResult
        {
            public string Method { get; set; } = "PATCH";
            public string Outcome { get; set; } // "INSERT" | "UPDATE" | "SUCCESS"
            public int StatusCode { get; set; }
            public string SalesforceId { get; set; } // em INSERT (201) pode vir no body
            public string RawBody { get; set; }      // resposta crua
        }

        internal class SalesforceUpsertResponse
        {
            public string id { get; set; }
            public bool success { get; set; }
            public bool created { get; set; }
            public object[] errors { get; set; }
        }
        // ---------------

        /// <summary>
        /// Upsert em OrderItem via External Id no PATH:
        /// PATCH {ApiCondicaoBase}/{ApiCondicaoExternalField}/{idExterno}
        /// - NÃO inclua o External Id no body.
        /// - Nulls são ignorados na serialização.
        /// </summary>
        public async Task<UpsertResult> UpsertItemPedido(string accessToken, string idExterno, object bodyObj)
        {
            var externalPath = string.Format("{0}/{1}/{2}",
                ConfigUrls.ApiCondicaoBase,
                ConfigUrls.ApiCondicaoExternalField,
                Uri.EscapeDataString(idExterno));

            var json = JsonConvert.SerializeObject(
                bodyObj,
                new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore }
            );

            var content = new StringContent(json, Encoding.UTF8, "application/json");

            using (var req = new HttpRequestMessage(new HttpMethod("PATCH"), externalPath))
            {
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                req.Content = content;

                using (var resp = await _http.SendAsync(req))
                {
                    var raw = await resp.Content.ReadAsStringAsync();
                    var result = new UpsertResult
                    {
                        Method = "PATCH",
                        StatusCode = (int)resp.StatusCode,
                        RawBody = raw
                    };

                    if (resp.IsSuccessStatusCode)
                    {
                        if (result.StatusCode == 201)
                        {
                            result.Outcome = "INSERT";
                            try
                            {
                                var parsed = JsonConvert.DeserializeObject<SalesforceUpsertResponse>(raw);
                                result.SalesforceId = parsed != null ? parsed.id : null;
                            }
                            catch { /* ignore */ }
                        }
                        else if (result.StatusCode == 204)
                        {
                            result.Outcome = "UPDATE";
                        }
                        else
                        {
                            result.Outcome = "SUCCESS";
                        }

                        return result;
                    }

                    throw new Exception(string.Format("Erro no UPSERT ItensPedidos (HTTP {0}): {1}", result.StatusCode, raw));
                }
            }
        }

        /// <summary>
        /// Executa um SOQL e retorna o primeiro registro como JObject (ou null).
        /// Usa a base REST derivada de ConfigUrls.ApiCondicaoBase (…/services/data/vXX.X).
        /// </summary>
        public async Task<JObject> QuerySingleAsync(string accessToken, string soql)
        {
            var restBase = GetRestBaseFromSObjectBase(); // ex.: https://.../services/data/v60.0
            var url = string.Format("{0}/query?q={1}", restBase, Uri.EscapeDataString(soql));

            using (var req = new HttpRequestMessage(HttpMethod.Get, url))
            {
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                using (var res = await _http.SendAsync(req))
                {
                    var raw = await res.Content.ReadAsStringAsync();
                    if (!res.IsSuccessStatusCode)
                        throw new Exception(string.Format("SOQL failed ({0}): {1}", (int)res.StatusCode, raw));

                    var root = JObject.Parse(raw);
                    var records = root["records"] as JArray;
                    return (records != null && records.Count > 0) ? (JObject)records[0] : null;
                }
            }
        }

        /// <summary>
        /// PATCH genérico em sObject: /sobjects/{sObjectName}/{id}
        /// Ex.: PatchSObject(token, "Order", orderId, new { Pricebook2Id = "01s..." })
        /// </summary>
        public async Task PatchSObject(string accessToken, string sObjectName, string id, object body)
        {
            var restBase = GetRestBaseFromSObjectBase();
            var url = string.Format("{0}/sobjects/{1}/{2}", restBase, sObjectName, id);

            var json = JsonConvert.SerializeObject(
                body,
                new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore }
            );
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            using (var req = new HttpRequestMessage(new HttpMethod("PATCH"), url))
            {
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                req.Content = content;

                using (var res = await _http.SendAsync(req))
                {
                    if (!res.IsSuccessStatusCode)
                    {
                        var raw = await res.Content.ReadAsStringAsync();
                        throw new Exception(string.Format("PATCH {0}/{1} falhou ({2}): {3}", sObjectName, id, (int)res.StatusCode, raw));
                    }
                }
            }
        }

        /// <summary>
        /// Deriva …/services/data/vXX.X a partir de ConfigUrls.ApiCondicaoBase (…/sobjects/OrderItem).
        /// </summary>
        private static string GetRestBaseFromSObjectBase()
        {
            // Ex.: https://xxx.my.salesforce.com/services/data/v60.0/sobjects/OrderItem
            var sobjectBase = ConfigUrls.ApiCondicaoBase;
            var idx = sobjectBase.IndexOf("/sobjects/", StringComparison.OrdinalIgnoreCase);
            if (idx <= 0)
                throw new Exception("ConfigUrls.ApiCondicaoBase inválida: não contém /sobjects/");

            // -> https://xxx.my.salesforce.com/services/data/v60.0
            return sobjectBase.Substring(0, idx);
        }
    }
}
