using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace StarkBankExcel
{
    internal class PaymentRequest
    {
        public static JObject Create(List<Dictionary<string, object>> payloads, string privateKeyPem = null)
        {
            Dictionary<string, object> body = new Dictionary<string, object>
            {
                { "requests", payloads }
            };

            return Request.Fetch(
                Request.Post,
                Globals.Credentials.Range["B3"].Value,
                "payment-request",
                body,
                null,
                privateKeyPem
            ).ToJson();
        }

        public static JObject Get(string cursor = null, Dictionary<string, object> optionalParams = null)
        {
            string query = "";

            if (cursor != null)
            {
                query = "?cursor=" + cursor;
            }

            if (optionalParams != null)
            {
                foreach (string key in optionalParams.Keys)
                {
                    if (query == "")
                    {
                        query = "?" + key + "=" + optionalParams[key].ToString();
                    }
                    else
                    {
                        query = query + "&" + key + "=" + optionalParams[key].ToString();
                    }
                }
            }

            return Request.Fetch(
                Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "payment-request" + query
            ).ToJson();
        }

    }
}
