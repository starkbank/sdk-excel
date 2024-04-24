using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using StarkBankExcel.Resources;

namespace StarkBankExcel
{
    public partial class SendReceiver
    {
        private void Planilha20_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha20_Shutdown(object sender, System.EventArgs e)
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
            this.Startup += new System.EventHandler(this.Planilha20_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha20_Shutdown);

        }

        #endregion

        private void button3_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.SendReceiver;

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Nome";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "CNPJ/CPF";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Banco";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Âgencia";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Conta";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Tipo";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "TAG";

            string returnMessage = "";
            string warningMessage = "";
            string errorMessage = "";

            int iteration = 0;

            var initRow = TableFormat.HeaderRow + 1;

            Dictionary<string, object> returnedData = new Dictionary<string, object>();
            List<Dictionary<string, object>> splits = new List<Dictionary<string, object>>();
            List<int> splitNumbers = new List<int>();

            for (int row = initRow; row <= lastRow; row++)
            {
                iteration++;

                string name = worksheet.Range["A" + row].Value?.ToString();
                string taxId = worksheet.Range["B" + row].Value?.ToString();
                string bank = worksheet.Range["C" + row].Value?.ToString(); // a validar
                string branchCode = worksheet.Range["D" + row].Value?.ToString();
                string accountNumber = worksheet.Range["E" + row].Value?.ToString();
                string type = worksheet.Range["F" + row].Value?.ToString();
                string tag = worksheet.Range["G" + row].Value?.ToString();

                Dictionary<string, object> split = new Dictionary<string, object>
                {
                    { "taxId", taxId },
                    { "name", name },
                    { "bankCode", bank },
                    { "accountNumber",  accountNumber },
                    { "branchCode", branchCode },
                    { "accountType", type },
                };

                if (tag != null)
                {
                    split.Add("tags", new List<string> { taxId, tag });
                };

                splits.Add(split);

                splitNumbers.Add(iteration);

                if (iteration % 100 == 0 || row >= lastRow)
                {
                    if (splitNumbers.Count == 0) goto nextIteration;

                    try
                    {

                        Dictionary<string, object> body = new Dictionary<string, object>
                        {
                            { "receivers", splits }
                        };

                        JObject res = Request.Fetch(
                            Request.Post,
                            Globals.Credentials.Range["B3"].Value,
                            "split-receiver",
                            body
                        ).ToJson();

                        string createSplit = res.ToString();

                        if (createSplit != null)
                        {
                            returnMessage = "Receivers criados com sucesso !" + "\n";
                        }
                    }
                    catch (Exception ex)
                    {
                        errorMessage = returnMessage + Utils.rowsMessage(initRow, row) + ex.Message + "\n";
                    }
                nextIteration:
                    initRow = row + 1;
                    splits = new List<Dictionary<string, object>>();
                    splitNumbers = new List<int>();
                }

                MessageBox.Show(warningMessage + returnMessage + errorMessage);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }
    }
}
