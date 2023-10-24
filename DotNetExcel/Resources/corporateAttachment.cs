using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StarkBankExcel
{
    internal class corporateAttachment
    {
        public static JObject Get(string key)
        {

            return V2Request.Fetch(
                V2Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "attachment/" + key + "?expand=content"
            ).ToJson();
        }
    }
}
