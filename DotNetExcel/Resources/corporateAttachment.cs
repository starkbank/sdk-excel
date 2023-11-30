using System;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace StarkBankExcel
{
    internal class corporateAttachment
    {
        public static JObject Get(string id, string cursor = null, Dictionary<string, object> optionalParams = null)
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
                V2Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "attachment/" + id + "/" + query
            ).ToJson();
        }

    }
}
