﻿using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Windows.Forms;

namespace StarkBankExcel.Resources
{
    internal class Split
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
                "split/" + query
            ).ToJson();
        }
    }
}
