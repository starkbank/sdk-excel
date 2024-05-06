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
            Environment.Items.Add("Produção");
            Environment.Items.Add("Ambiente Desenvolvedor");
            // Environment.Items.Add("Development");

            Environment.Text = "Produção";
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            string environment = Environment.Text.ToLower();
            string workspace = Workspace.Text.ToLower();
            string email = Email.Text.ToLower();
            string password = Password.Text.ToString();

            if (environment.Trim().Length == 8)
            {
                environment = "production";
            }

            if (environment.Trim().Length == 22)
            {
                environment = "sandbox";
            }

            if (environment.Trim().Length == 11)
            {
                environment = "development";
            }

            try
            {
                Session.Create(workspace, environment, email, password);
            }
            catch (Exception ex)
            {
                MessageBox.Show("O workspace " + workspace.Trim() + " não existe no ambiente selecionado");
                return;
            }

            PrivateKey privateKey = new PrivateKey();
            PublicKey publicKey = privateKey.publicKey();

            List<object> challenge = new List<object>();

            Dictionary<string, object> requestBody = new Dictionary<string, object>()
            {
                { "platform", "web" },
                { "expiration", 604800 },
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

            int loginInt = 0;

            try
            {

                fetchedJson = Request.Fetch(
                     Request.Post,
                     environment,
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

                qrCode formSize = new qrCode();

                Globals.Credentials.Range["C16"].Value = formSize.Size.ToString();

                formSize.Close();

                qrCode qrcodeForms = new qrCode();

                Task qrcodeTask = Task.Run(() => qrcodeWork(qrcodeForms));
                Task<int> pollingTask = Task.Run(() => pollingWork());

                Task.WaitAny(qrcodeTask, pollingTask);

                if (await pollingTask == 1)
                {
                    loginInt = 1;

                    Task closeQrcodeTask = Task.Run(() => closeQrcodeWork(qrcodeForms));
                    Task.WaitAny(closeQrcodeTask);

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
                    Globals.Credentials.Range["B13"].Value = "session/" + fetchedJson.ToJson()["session"]["id"].ToString();


                    Workbook workbook = Globals.ThisWorkbook.Application.ActiveWorkbook;

                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        if (sheet.Name != "Credentials")
                        {
                            Utils.DisplayInfo(sheet);
                        }
                    }

                }

                if (await pollingTask == 2)
                {
                    loginInt = 2;

                    qrcodeForms.DialogResult = DialogResult.Cancel;
                    qrcodeForms.Close();
                }

                if (await pollingTask == 3)
                {
                    loginInt = 3;

                    qrcodeForms.DialogResult = DialogResult.Cancel;
                    qrcodeForms.Close();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }

            if (loginInt == 1)
            {
                MessageBox.Show("Logado com sucesso!");
                Close();
            }
            if (loginInt == 2)
            {
                MessageBox.Show("Acesso Negado!");
                Close();
            }
            if (loginInt == 3)
            {
                MessageBox.Show("Este QR Code expirou. Efetue o login novamente para gerar um novo QR Code antes de continuar.");
            }
        }

        static void qrcodeWork(Form qrcodeForms)
        {
            
            qrcodeForms.ShowDialog();

        }

        static int pollingWork()
        {
            string id = Globals.Credentials.Range["B15"].Value;
            string challengePk = Globals.Credentials.Range["B16"].Value;
            string development = Globals.Credentials.Range["B3"].Value;


            string path = "challenge/" + id;

            for (int i = 0; i < 55; i++)
            {
                Response response = Request.Fetch(
                  Request.Get,
                  development,
                  path,
                  null,
                  null,
                  challengePk
                );

                if (response.ToJson()["challenge"]["status"].ToString() == "approved")
                {
                    return 1;
                }

                if (response.ToJson()["challenge"]["status"].ToString() == "denied")
                {
                    return 2;
                } 

                if (response.ToJson()["challenge"]["status"].ToString() == "expired")
                {
                    return 3;
                }
                

                Thread.Sleep(1000);
            }

            return 3;
        }

        static void closeQrcodeWork(Form qrcodeForms)
        {
            qrcodeForms.DialogResult = DialogResult.Cancel;
            qrcodeForms.Close();
        }

    }
}
