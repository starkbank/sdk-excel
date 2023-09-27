using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace StarkBankExcel.Resources
{
    internal class Boleto
    {
        public static JObject Create(List<Dictionary<string, object>> payloads)
        {
            Dictionary<string, object> body = new Dictionary<string, object>
            {
                { "boletos", payloads }
            };

            return V2Request.Fetch(
                V2Request.Post,
                Globals.Credentials.Range["B3"].Value,
                "boleto",
                body
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

            return V2Request.Fetch(
                V2Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "boleto/" + query
            ).ToJson();
        }

        public class Log
        {
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

                return V2Request.Fetch(
                    V2Request.Get,
                    Globals.Credentials.Range["B3"].Value,
                    "boleto/log/" + query
                ).ToJson();
            }
        }
    }
}
