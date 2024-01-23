using Microsoft.VisualStudio.Tools.Applications.Runtime;
using StarkBankExcel.Forms;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace StarkBankExcel
{
    public partial class GetPaymentApprove
    {
        private void Planilha15_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha15_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button3.Click += new System.EventHandler(this.button3_Click);
            this.button4.Click += new System.EventHandler(this.button4_Click);
            this.button5.Click += new System.EventHandler(this.button5_Click);
            this.Startup += new System.EventHandler(this.Planilha15_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha15_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            ViewPaymentApprove viewPaymentApprovement = new ViewPaymentApprove();
            viewPaymentApprovement.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Globals.Main.Activate();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            LoginForm loginForm = new LoginForm();
            loginForm.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Utils.LogOut();
        }
    }
}
