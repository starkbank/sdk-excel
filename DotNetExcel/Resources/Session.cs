using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace StarkBankExcel.Resources
{
    internal class Session
    {
        public static void Create(string workspace, string environment, string email, string password)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>()
            {
                { "workspace", workspace },
                { "email", email },
                { "password", password },
                { "platform", "web" }
            };

            JObject fetchedJson;

            try
            {
                fetchedJson = Request.Fetch(
                    Request.Post,
                    environment,
                    "auth/access-token",
                    payload
                ).ToJson();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            SaveSession(workspace, environment, email, fetchedJson["accessToken"].ToString(), 
                        fetchedJson["member"]["name"].ToString(), fetchedJson["member"]["workspaceId"].ToString());
        }

        public static void Delete()
        {
            JObject fetchedJson;
            try
            {
                fetchedJson = Request.Fetch(
                    Request.Delete,
                    Globals.Credentials.Range["B3"].Value,
                    "auth/access-token/" + Globals.Credentials.Range["B4"].Value
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
