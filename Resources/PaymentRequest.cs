using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace StarkBankExcel
{
    internal class PaymentRequest
    {
        public static JObject Create(List<Dictionary<string, object>> payloads)
        {
            Dictionary<string, object> body = new Dictionary<string, object>
            {
                { "requests", payloads }
            };

            return V2Request.Fetch(
                V2Request.Post,
                Globals.Credentials.Range["B3"].Value,
                "payment-request",
                body
            ).ToJson();
        }
    }
}
