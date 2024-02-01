using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StarkBankExcel.Resources
{
    internal class BoletoPaymentClass
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
                "boleto-payment" + query
            ).ToJson();
        }

        public static JObject Create(Dictionary<string, object> payload = null, string privateKey = null)
        {
            return Request.Fetch(
                Request.Post,
                Globals.Credentials.Range["B3"].Value,
                "boleto-payment",
                payload,
                null,
                privateKey
            ).ToJson();
        }

    }
}
