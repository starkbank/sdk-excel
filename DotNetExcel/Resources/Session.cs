using System;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace StarkBankExcel.Resources
{
    internal class Session
    {
        public static void Create(string workspace, string environment, string email, string password)
        {
            HttpRequestMessage httpRequestMessage = new HttpRequestMessage
            {
                Method = new HttpMethod("GET"),
                RequestUri = new Uri(Utils.BaseUrl(environment) + "v2/workspace?username=" + workspace)
            };

            HttpClient Client = new HttpClient();
            Client.DefaultRequestHeaders.Add("User-Agent", "Excel-DotNet");
            httpRequestMessage.Headers.TryAddWithoutValidation("Content-Type", "application/json");
            httpRequestMessage.Headers.TryAddWithoutValidation("Accept-Language", "pt-BR");

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            var result = Client.SendAsync(httpRequestMessage).Result;

            Response response = new Response(
                result.Content.ReadAsByteArrayAsync().Result,
                (int)result.StatusCode
                );

            var workSpaceId = response.ToJson()["workspaces"][0]["id"];
            var memberName = response.ToJson()["workspaces"][0]["username"];

            SaveSession(workspace, environment, email, "", 
                        memberName.ToString(), workSpaceId.ToString());
        }

        public static void Delete()
        {
            JObject fetchedJson;
            try
            {
                fetchedJson = V2Request.Fetch(
                    Request.Delete,
                    Globals.Credentials.Range["B3"].Value,
                    "session/" + Globals.Credentials.Range["B13"].Value.Split("/")[1]
               ).ToJson();
            } catch (Exception) { }
            
            SaveSession("", "", "", "", "", "");
        }

        private static void SaveSession(string workspace, string environment, string email, string accessToken, string name, string workspaceId)
        {
            Globals.Credentials.Range["B1"].Value = workspace;
            Globals.Credentials.Range["B2"].Value = email;
            Globals.Credentials.Range["B3"].Value = environment;
            Globals.Credentials.Range["B4"].Value = accessToken;
            Globals.Credentials.Range["B5"].Value = name;
            Globals.Credentials.Range["B6"].Value = workspaceId;
        }
    }
}