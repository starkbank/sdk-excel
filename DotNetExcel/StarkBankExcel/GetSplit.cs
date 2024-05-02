using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using StarkBankExcel.Resources;
using StarkBankExcel.Forms;

namespace StarkBankExcel
{
    public partial class GetSplit
    {
        private void Planilha19_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha19_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.button3.Click += new System.EventHandler(this.button3_Click);
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button5.Click += new System.EventHandler(this.button5_Click);
            this.button6.Click += new System.EventHandler(this.button6_Click);
            this.Startup += new System.EventHandler(this.Planilha19_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha19_Shutdown);

        }

        #endregion

        private void button3_Click(object sender, EventArgs e)
        {
            ViewSplit viewSplit = new ViewSplit();
            viewSplit.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Globals.Main.Activate();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            LoginForm loginForm = new LoginForm();
            loginForm.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Utils.LogOut();
        }
    }
}
