using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace StarkBankExcel
{
    public partial class SendCartShop
    {
        private void Planilha18_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha18_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Planilha18_Startup);
            this.Shutdown += new System.EventHandler(Planilha18_Shutdown);
        }

        #endregion

    }
}
