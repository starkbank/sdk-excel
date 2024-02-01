using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.IO;

namespace StarkBankExcel
{
    internal class Json
    {
        internal static string Encode(object payload)
        {
            return JsonConvert.SerializeObject(payload);
        }

        internal static JObject Decode(string content)
        {
            using (var reader = new JsonTextReader(new StringReader(content)) { DateParseHandling = DateParseHandling.None })
                return JObject.Load(reader);
        }
    }
}
