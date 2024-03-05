using System;
using System.Data;
using System.Drawing;
using StarkBankExcel.Forms;
using System.Windows.Forms;
using StarkBankExcel.Resources;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System.Net.Http;
using System.Net;

namespace StarkBankExcel
{
    public partial class Main
    {
        private void Planilha1_Startup(object sender, System.EventArgs e)
        {

            var worksheet = Globals.Main;
            string version = worksheet.Range["A1"].Value;

            version = version.ToString().Split('v')[1].Trim();

            string url = "https://github.com/starkbank/sdk-excel/blob/master/CHANGELOG.md";

            HttpRequestMessage httpRequestMessage = new HttpRequestMessage
            {
                Method = new HttpMethod("GET"),
                RequestUri = new Uri(url)
            };

            HttpClient Client = new HttpClient();
            Client.DefaultRequestHeaders.Add("User-Agent", "Excel-DotNet");
            httpRequestMessage.Headers.TryAddWithoutValidation("Content-Type", "application/json");
            httpRequestMessage.Headers.TryAddWithoutValidation("Accept-Language", "pt-BR");
            httpRequestMessage.Headers.TryAddWithoutValidation("Accept", "*/*");

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            var result = Client.SendAsync(httpRequestMessage).Result;

            Response response = new Response(
                result.Content.ReadAsByteArrayAsync().Result,
                (int)result.StatusCode
                );

            var versionWarning = response.ToJson()["payload"]["blob"]["headerInfo"]["toc"][1]["text"];

            versionWarning = versionWarning.ToString().Split(']')[0].Split('[')[1];

            if (version.ToString().Trim() != versionWarning.ToString().Trim())
            {
                VersionWarning viewInvoiceForm = new VersionWarning();
                viewInvoiceForm.ShowDialog();
            }
        }

        private void Planilha1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.login.Click += new System.EventHandler(this.button1_Click);
            this.button2.Click += new System.EventHandler(this.button2_Click);
            this.transferOrder.Click += new System.EventHandler(this.transferOrder_Click);
            this.getDictKey.Click += new System.EventHandler(this.getDictKey_Click);
            this.Invoice.Click += new System.EventHandler(this.Invoice_Click);
            this.button7.Click += new System.EventHandler(this.button7_Click);
            this.Help.Click += new System.EventHandler(this.Help_Click);
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            this.button8.Click += new System.EventHandler(this.button8_Click);
            this.button9.Click += new System.EventHandler(this.button9_Click);
            this.button10.Click += new System.EventHandler(this.button10_Click);
            this.button11.Click += new System.EventHandler(this.button11_Click);
            this.button12.Click += new System.EventHandler(this.button12_Click);
            this.Startup += new System.EventHandler(this.Planilha1_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha1_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            LoginForm loginForm = new LoginForm();
            loginForm.ShowDialog();
        }

        private void transferOrder_Click(object sender, EventArgs e)
        {
            Globals.Transfers.Activate();
        }

        private void getDictKey_Click(object sender, EventArgs e)
        {
            Globals.GetDictKeys.Activate();
        }

        private void Invoice_Click(object sender, EventArgs e)
        {
            Globals.GetInvoices.Activate();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Globals.SendInvoices.Activate();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Utils.LogOut();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Globals.GetStatement.Activate();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Globals.GetBoleto.Activate();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Globals.GetBoletoEvents.Activate();
        }

        private void Help_Click(object sender, EventArgs e)
        {
            ViewHelpForm viewHelpForm = new ViewHelpForm();
            viewHelpForm.ShowDialog();

        }

        private void button10_Click(object sender, EventArgs e)
        {
            Globals.Planilha12.Activate();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Globals.Planilha11.Activate();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            Globals.GetPaymentApprove.Activate();
        }
    }
}
