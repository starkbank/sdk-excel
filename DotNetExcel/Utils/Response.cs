using Newtonsoft.Json.Linq;
using StarkBankExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StarkBankExcel
{
    internal class Response
    {
        internal byte[] ByteContent { get; }
        internal int Status { get; }

        internal Response(byte[] byteContent, int status)
        {
            ByteContent = byteContent;
            Status = status;
        }

        internal string Content
        {
            get
            {
                return System.Text.Encoding.UTF8.GetString(ByteContent);
            }
        }

        internal JObject ToJson()
        {
            return Json.Decode(Content);
        }
    }
}
