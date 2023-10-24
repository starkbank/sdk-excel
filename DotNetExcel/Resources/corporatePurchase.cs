using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace StarkBankExcel.Resources
{
    internal class corporatePurchase
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

            //MessageBox.Show(query);

            return V2Request.Fetch(
                V2Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "corporate-purchase" + query
            ).ToJson();
        }
    }
}
