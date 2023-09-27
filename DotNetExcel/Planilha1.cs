using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace StarkBankMVP
{
    public partial class Planilha1
    {
        private void Planilha1_Startup(object sender, System.EventArgs e)
        {
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
            this.button3.Click += new System.EventHandler(this.button3_Click);
            this.transferOrder.Click += new System.EventHandler(this.transferOrder_Click);
            this.getDictKey.Click += new System.EventHandler(this.getDictKey_Click);
            this.Invoice.Click += new System.EventHandler(this.Invoice_Click);
            this.button7.Click += new System.EventHandler(this.button7_Click);
            this.Startup += new System.EventHandler(this.Planilha1_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha1_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            LoginForm loginForm = new LoginForm();
            loginForm.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {

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
            Globals.CreateInvoices.Activate();
        }
    }
}
