using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace StarkBankExcel
{
    public partial class SendBoleto
    {
        private void Planilha10_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha10_Shutdown(object sender, System.EventArgs e)
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
            this.button6.Click += new System.EventHandler(this.button6_Click);
            this.Startup += new System.EventHandler(this.Planilha10_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha10_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.SendBoleto;

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":S" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Nome";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "CPF/CNPJ";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Logradouro";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Complemento";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Bairo";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Cidade";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Código do Estado";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "CEP";
            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["J" + TableFormat.HeaderRow].Value = "Data de Vencimento";
            worksheet.Range["K" + TableFormat.HeaderRow].Value = "Multa";
            worksheet.Range["L" + TableFormat.HeaderRow].Value = "Juro ao Mês";
            worksheet.Range["M" + TableFormat.HeaderRow].Value = "Dias para Baixa Automática";
            worksheet.Range["N" + TableFormat.HeaderRow].Value = "Descrição 1";
            worksheet.Range["O" + TableFormat.HeaderRow].Value = "Valor 1";
            worksheet.Range["P" + TableFormat.HeaderRow].Value = "Descrição 2";
            worksheet.Range["Q" + TableFormat.HeaderRow].Value = "Valor 2";
            worksheet.Range["R" + TableFormat.HeaderRow].Value = "Descrição 3";
            worksheet.Range["S" + TableFormat.HeaderRow].Value = "Valor 3";

            string returnMessage = "";
            string warningMessage = "";
            string errorMessage = "";

            int iteration = 0;

            var initRow = TableFormat.HeaderRow + 1;
            lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            List<Dictionary<string, object>> boletos = new List<Dictionary<string, object>>();
            List<int> boletoNumbers = new List<int>();

            for (int row = initRow; row <= lastRow; row++)
            {
                iteration++;

                string name = worksheet.Range["A" + row].Value?.ToString();
                string taxID = worksheet.Range["B" + row].Value?.ToString();
                string amountString = worksheet.Range["C" + row].Value?.ToString();
                int amount = int.Parse(amountString);
                string due = worksheet.Range["D" + row].Value?.ToString();
                string fineString = worksheet.Range["E" + row].Value?.ToString();
                string interestString = worksheet.Range["F" + row].Value?.ToString().Replace(",", ".");
                string expirationString = worksheet.Range["G" + row].Value?.ToString().Replace(",", ".");

                List<Dictionary<string, string>> descriptions = new List<Dictionary<string, string>>();

                string description1 = worksheet.Range["H" + row].Value?.ToString();
                string value1 = worksheet.Range["I" + row].Value?.ToString();
                string description2 = worksheet.Range["J" + row].Value?.ToString();
                string value2 = worksheet.Range["K" + row].Value?.ToString();
                string description3 = worksheet.Range["L" + row].Value?.ToString();
                string value3 = worksheet.Range["M" + row].Value?.ToString();

                if (description1 != null && value1 != null)
                {
                    descriptions.Add(new Dictionary<string, string>
                    {
                        {"key", description1 },
                        {"value", value1 },
                    });
                }

                if (description2 != null && value2 != null)
                {
                    descriptions.Add(new Dictionary<string, string>
                    {
                        {"key", description2 },
                        {"value", value2 },
                    });
                }

                if (description3 != null && value3 != null)
                {
                    descriptions.Add(new Dictionary<string, string>
                    {
                        {"key", description3 },
                        {"value", value3 },
                    });
                }

                Dictionary<string, object> boleto = new Dictionary<string, object> {
                    {"amount", amount },
                    {"taxId", taxID },
                    {"name", name},
                    {"descriptions" , descriptions }
                };

                if (due != null) boleto.Add("due", new StarkDateTime(due).ToString());
                if (expirationString != null) boleto.Add("expiration", int.Parse(expirationString));
                if (fineString != null) boleto.Add("fine", float.Parse(fineString));
                if (interestString != null) boleto.Add("interest", float.Parse(interestString));

                boletoNumbers.Add(iteration);

                boletos.Add(boleto);

                if (iteration % 100 == 0 || row >= lastRow)
                {
                    if (boletoNumbers.Count == 0) goto nextIteration;

                    try
                    {
                        JObject res = Boleto.Create(boletos);
                        string createBoleto = (string)res["message"];
                        returnMessage = returnMessage + Utils.rowsMessage(initRow, row) + createBoleto + "\n";
                    }
                    catch (Exception ex)
                    {
                        errorMessage = ex.Message;
                    }
                nextIteration:
                    initRow = row + 1;
                    boletos = new List<Dictionary<string, object>>();
                    boletoNumbers = new List<int>();
                }
            }

            MessageBox.Show(warningMessage + returnMessage + errorMessage);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.SendBoleto;

            Range range = worksheet.Range["A" + (TableFormat.HeaderRow + 1) + ":S1048576"];
            range.ClearContents();
        }
    }
}
