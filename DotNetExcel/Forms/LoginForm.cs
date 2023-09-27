using EllipticCurve;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace StarkBankMVP
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        private void Login_Click(object sender, EventArgs e)
        {
            string environment = Environment.Text.ToLower();
            string workspace = Workspace.Text.ToLower();
            string email = Email.Text.ToLower();
            string password = Password.Text.ToString();

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
                MessageBox.Show(ex.ToString());
                Close();
                return;
            }

            Globals.Credentials.Range["B1"].Value = workspace;
            Globals.Credentials.Range["B2"].Value = email;
            Globals.Credentials.Range["B3"].Value = environment;
            Globals.Credentials.Range["B4"].Value = fetchedJson["accessToken"].ToString();
            Globals.Credentials.Range["B5"].Value = fetchedJson["member"]["name"].ToString();
            Globals.Credentials.Range["B6"].Value = fetchedJson["member"]["workspaceId"].ToString();

            PrivateKey privateKey = new PrivateKey();
            PublicKey publicKey = privateKey.publicKey();

            Dictionary<string, object> teste = new Dictionary<string, object>()
            {
                { "platform", "web" },
                { "expiration", 5184000 },
                { "publicKey", publicKey.toPem() }
            };

            JObject fetchedJson2;

            try
            {
                fetchedJson2 = Request.Fetch(
                    Request.Post,
                    environment,
                    "auth/session",
                    teste
                ).ToJson();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                Close();
                return;
            }

            Globals.Credentials.Range["A11"].Value = "Session Private";
            Globals.Credentials.Range["B11"].Value = privateKey.toPem();

            Globals.Credentials.Range["A12"].Value = "Session Public";
            Globals.Credentials.Range["B12"].Value = publicKey.toPem();

            Globals.Credentials.Range["A13"].Value = "Access ID";
            Globals.Credentials.Range["B13"].Value = "session/" + fetchedJson2["session"]["id"].ToString();

            MessageBox.Show("Logado com sucesso!");
            Close();
        }

        private void LoginForm_Load(object sender, EventArgs e)
        {
            Environment.Items.Add("Produção");
            Environment.Items.Add("Sandbox");

            Environment.Text = "Produção";
        }
    }
}
