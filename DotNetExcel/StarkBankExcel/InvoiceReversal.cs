using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace StarkBankExcel
{
    public partial class InvoiceReversal
    {
        private void Planilha13_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha13_Shutdown(object sender, System.EventArgs e)
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
            this.button4.Click += new System.EventHandler(this.button4_Click);
            this.button5.Click += new System.EventHandler(this.button5_Click);
            this.button6.Click += new System.EventHandler(this.button6_Click);
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.Startup += new System.EventHandler(this.Planilha13_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha13_Shutdown);

        }

        #endregion

        private void button6_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.InvoiceReversal;

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "ID da Invoice";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "Valor Final";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Status";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Erro";

            string returnMessage = "";
            string warningMessage = "";
            string errorMessage = "";

            int iteration = 0;

            var initRow = TableFormat.HeaderRow + 1;
            lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            List<int> invoiceReversalNumbers = new List<int>();

            for (int row = initRow; row <= lastRow; row++)
            {
                iteration++;

                string invoiceId = worksheet.Range["A" + row].Value?.ToString();
                string amountString = worksheet.Range["B" + row].Value?.ToString();
                int amount = Convert.ToInt32(double.Parse(amountString) * 100);

                Dictionary<string, object> invoiceReversal = new Dictionary<string, object> {
                    { "amount", amount }
                };

                invoiceReversalNumbers.Add(iteration);

                if (invoiceReversalNumbers.Count == 0) goto nextIteration;

                try
                {
                    JObject res = Invoice.Update(invoiceId, invoiceReversal);

                    string createInvoice = (string)res["message"];
                    worksheet.Range["C" + row].Value = createInvoice;
                    worksheet.Range["D" + row].Value = "";
                    returnMessage = returnMessage + Utils.rowsMessage(initRow, row) + createInvoice + "\n";
                }
                catch (Exception ex)
                {
                    errorMessage = ex.Message;
                    worksheet.Range["C" + row].Value = "";
                    worksheet.Range["D" + row].Value = ex.Message;
                }
            nextIteration:
                initRow = row + 1;
                invoiceReversalNumbers = new List<int>();
            }

            MessageBox.Show(warningMessage + returnMessage + errorMessage);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.InvoiceReversal;

            Range range = worksheet.Range["A" + (TableFormat.HeaderRow + 1) + ":K1048576"];
            range.ClearContents();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Globals.Main.Activate();
        }

        private void button4_Click(object sender, EventArgs e)
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
