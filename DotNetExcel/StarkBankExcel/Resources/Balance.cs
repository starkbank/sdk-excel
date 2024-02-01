using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace StarkBankExcel.Resources
{
    internal class Balance
    {
        public static double Get()
        {
            JObject respJson = Request.Fetch(
                Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "balance"
            ).ToJson();

            return (double)respJson["balances"][0]["amount"];
        }
    }
}
