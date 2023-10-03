using EllipticCurve;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace StarkBankExcel
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

            try
            {
                Session.Create(workspace, environment, email, password);
            } 
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }

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
                return;
            }

            Globals.Credentials.Range["A11"].Value = "Session Private";
            Globals.Credentials.Range["B11"].Value = privateKey.toPem();

            Globals.Credentials.Range["A12"].Value = "Session Public";
            Globals.Credentials.Range["B12"].Value = publicKey.toPem();

            Globals.Credentials.Range["A13"].Value = "Access ID";
            Globals.Credentials.Range["B13"].Value = "session/" + fetchedJson2["session"]["id"].ToString();


            Workbook workbook = Globals.ThisWorkbook.Application.ActiveWorkbook;

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if(sheet.Name != "Credentials")
                {
                    Utils.DisplayInfo(sheet);
                }
            }
            
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
