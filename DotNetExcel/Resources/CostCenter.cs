using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace StarkBankExcel
{
    internal class CostCenter
    {
        public static JObject Get(string cursor = null, Dictionary<string, object> optionalParams = null)
        {
            string query = "";

            if(cursor != null)
            {
                query = "?cursor=" + cursor;    
            }

            if(optionalParams != null)
            {
                foreach(string key in optionalParams.Keys)
                {
                    if(query == "")
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
                "cost-center/" + query
            ).ToJson();
        }
    }
}
