using System;
using System.Text;
using System.Linq;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace StarkBankExcel
{
    internal class DictKey
    {
        public static JObject Get(string key)
        {
            return Request.Fetch(
                V2Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "dict-key/" + key
            ).ToJson();
        }
    }
}
