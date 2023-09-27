using EllipticCurve;
using StarkBankMVP;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StarkBankMVP
{
    internal class V2Request
    {
        private static HttpClient makeClient()
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("User-Agent", "Excel-DotNet");
            return client;
        }

        private static readonly HttpClient Client = makeClient();
        internal static readonly HttpMethod Get = new HttpMethod("GET");
        internal static readonly HttpMethod Put = new HttpMethod("PUT");
        internal static readonly HttpMethod Post = new HttpMethod("POST");
        internal static readonly HttpMethod Patch = new HttpMethod("PATCH");
        internal static readonly HttpMethod Delete = new HttpMethod("DELETE");

        internal static Response Fetch(
            HttpMethod method, string environment, string path, Dictionary<string, object> payload = null,
            Dictionary<string, object> query = null
        )
        {
            string url = "";
            if (environment == "production")
            {
                url = "https://api.starkbank.com/";
            }
            if (environment == "sandbox")
            {
                url = "https://sandbox.api.starkbank.com/";
            }
            url += "v2/" + path;

            if (query != null)
            {
                url += Url.Encode(query);
            }

            string accessTime = DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalSeconds.ToString(new CultureInfo("en-US"));
            string accessId = Globals.Credentials.Range["B13"].Value;
            string body = "";

            if (payload != null)
            {
                body = Json.Encode(payload);
            }

            string privateKeyPem = Globals.Credentials.Range["B11"].Value;

            if(privateKeyPem == null)
            {
                throw new Exception("Credenciais Inválidas, necessário realizar Login novamente!");
            }
            PrivateKey privateKey = PrivateKey.fromPem(privateKeyPem);

            string message = accessId + ":" + accessTime + ":" + body;
            string signature = Ecdsa.sign(message, privateKey).toBase64();

            HttpRequestMessage httpRequestMessage = new HttpRequestMessage
            {
                Method = method,
                RequestUri = new Uri(url)
            };
            if (body.Length > 0)
            {
                httpRequestMessage.Content = new StringContent(body);
            }

            httpRequestMessage.Headers.TryAddWithoutValidation("Access-Id", accessId);
            httpRequestMessage.Headers.TryAddWithoutValidation("Access-Time", accessTime);
            httpRequestMessage.Headers.TryAddWithoutValidation("Access-Signature", signature);
            httpRequestMessage.Headers.TryAddWithoutValidation("Content-Type", "application/json");
            httpRequestMessage.Headers.TryAddWithoutValidation("Accept-Language", "pt-BR");

            var result = Client.SendAsync(httpRequestMessage).Result;

            Response response = new Response(
                result.Content.ReadAsByteArrayAsync().Result,
                (int)result.StatusCode
            );

            if (response.Status == 500)
            {
                MessageBox.Show("Internal Server Error \n Houston, we have a problem!");
            }
            if (response.Status == 400)
            {
                throw new Exception(response.Content);
            }
            if (response.Status != 200)
            {
                throw new Exception(response.Content);
            }

            return response;
        }
    }
}
