using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace StarkBankExcel.Resources
{
    internal class Balance
    {
        public static double Get()
        {
            JObject respJson = V2Request.Fetch(
                V2Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "balance"
            ).ToJson();

            return (double)respJson["balances"][0]["amount"];
        }
    }
}
