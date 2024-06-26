﻿using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace StarkBankExcel.Resources
{
    internal class SplitReceiver
    {
        public static JObject Create(List<Dictionary<string, object>> payloads)
        {
            Dictionary<string, object> body = new Dictionary<string, object>
            {
                { "receivers", payloads }
            };

            return Request.Fetch(
                Request.Post,
                Globals.Credentials.Range["B3"].Value,
                "split-receiver",
                body
            ).ToJson();
        }

        public static JObject Get(string cursor = null, Dictionary<string, object> optionalParams = null)
        {
            string query = "";

            if (cursor != null)
            {
                query = "?fields=id,taxId&cursor=" + cursor;
            }
            else
            {
                query = "?fields=id,taxId";
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
                "split-receiver/" + query
            ).ToJson();
        }
    }
}
