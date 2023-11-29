using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace StarkBankExcel.Resources
{
    internal class CorporateTransaction
    {

        public static JObject Get(string cursor = null, string optionalParams = null)
        {
            string query = optionalParams;

            if (cursor != null)
            {
                query = "?cursor=" + cursor;
            }

            return Request.Fetch(
                V2Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "corporate-transaction" + query
            ).ToJson();
        }
    }
}
