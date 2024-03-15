using System;
using EllipticCurve;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using Newtonsoft.Json;
using StarkBankExcel.Forms;
using Microsoft.Office.Tools.Excel.Controls;
using System.Threading;
using System.Threading.Tasks;

namespace StarkBankExcel
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        private void LoginForm_Load(object sender, EventArgs e)
        {
            Environment.Items.Add("Production");
            Environment.Items.Add("Sandbox");

            Environment.Text = "Production";
        }

        private void button1_Click(object sender, EventArgs e)
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

            List<object> challenge = new List<object>();

            Dictionary<string, object> requestBody = new Dictionary<string, object>()
            {
                { "platform", "web" },
                { "expiration", 5184000 },
                { "publicKey", publicKey.toPem() }
            };

            Dictionary<string, object> dictObj = new Dictionary<string, object>()
            {
                { "requestBody", JsonConvert.SerializeObject(requestBody) },
                { "requestMethod", "POST" },
                { "requestPath", "/session" },
                { "type", "authenticator" }
            };

            challenge.Add(dictObj);

            Dictionary<string, object> payload = new Dictionary<string, object>() { { "challenges", challenge } };

            Response fetchedJson;

            PrivateKey keys = keyGen.generateKeyFromPassword(password, email);

            try
            {

                fetchedJson = Request.Fetch(
                     Request.Post,
                     "sandbox",
                     "challenge?expand=qrcode",
                     payload,
                     null,
                     keys.toPem()
                  );

                string qrcode = fetchedJson.ToJson()["challenges"][0]["qrcode"].ToString();
                string challengeId = fetchedJson.ToJson()["challenges"][0]["id"].ToString();
                string challengPk = keys.toPem();

                Globals.Credentials.Range["A14"].Value = "Qrcode";
                Globals.Credentials.Range["B14"].Value = qrcode;

                Globals.Credentials.Range["A15"].Value = "challenge";
                Globals.Credentials.Range["B15"].Value = challengeId;

                Globals.Credentials.Range["A16"].Value = "ChallengePk";
                Globals.Credentials.Range["B16"].Value = challengPk;

                Task task1 = Task.Run(() => Dowork1());
                Task task2 = Task.Run(() => Dowork2());

                Task.WaitAll(task2);

                string dictObj2 = JsonConvert.SerializeObject(requestBody);

                try
                {
                    fetchedJson = Request.maskFetch(
                         Request.Post,
                         environment,
                         "/session",
                         dictObj2,
                         null,
                         keys.toPem(),
                         challengeId
                     );

                } catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    return;
                }

                Globals.Credentials.Range["A11"].Value = "Session Private";
                Globals.Credentials.Range["B11"].Value = privateKey.toPem();

                Globals.Credentials.Range["A12"].Value = "Session Public";
                Globals.Credentials.Range["B12"].Value = publicKey.toPem();

                Globals.Credentials.Range["A13"].Value = "Access ID";
                Globals.Credentials.Range["B13"].Value = "session/" + fetchedJson.ToJson()["session"]["id"].ToString();


                Workbook workbook = Globals.ThisWorkbook.Application.ActiveWorkbook;

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name != "Credentials")
                    {
                        Utils.DisplayInfo(sheet);
                    }
                }

                MessageBox.Show("Logado com sucesso!");

                Close();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }

            Close();
        }

        static void Dowork1()
        {
            qrCode qrcodeForms = new qrCode();
            qrcodeForms.ShowDialog();
        }

        static void Dowork2()
        {
            string id = Globals.Credentials.Range["B15"].Value;
            string challengePk = Globals.Credentials.Range["B16"].Value;


            string path = "challenge/" + id;

            for (int i = 0; i < 1000; i++)
            {
                Response response = Request.Fetch(
                  Request.Get,
                  "sandbox",
                  path,
                  null,
                  null,
                  challengePk
               );

                if (response.ToJson()["challenge"]["status"].ToString() == "approved")
                {
                    break;
                }

                Thread.Sleep(1000);
            }
        }

    }
}
