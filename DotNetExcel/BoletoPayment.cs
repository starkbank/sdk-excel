using System;
using System.Data;
using System.Drawing;
using StarkBankExcel.Forms;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace StarkBankExcel
{
    public partial class BoletoPayment
    {
        private void Planilha16_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha16_Shutdown(object sender, System.EventArgs e)
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
            this.Startup += new System.EventHandler(this.Planilha16_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha16_Shutdown);
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            BoletoPaymentForm boletoPayment = new BoletoPaymentForm();
            boletoPayment.ShowDialog();
        }
    }
}
