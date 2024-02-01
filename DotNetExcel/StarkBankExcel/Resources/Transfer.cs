using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using System;
using System.Text;

namespace StarkBankExcel.Resources
{
    internal class Transfer
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
            return Request.Fetch(
                Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "transfer/" + query
            ).ToJson();
        }

        public static byte[] Pdf(string id)
        {
            return Encoding.UTF8.GetBytes(Request.Fetch(
                Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "transfer/" + id + "/pdf"
            ).Content);
        }
    }
}
