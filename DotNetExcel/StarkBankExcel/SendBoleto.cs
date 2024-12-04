using System;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

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
            this.button3.Click += new System.EventHandler(this.button3_Click);
            this.button4.Click += new System.EventHandler(this.button4_Click);
            this.button5.Click += new System.EventHandler(this.button5_Click);
            this.button6.Click += new System.EventHandler(this.button6_Click_1);
            this.Startup += new System.EventHandler(this.Planilha10_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha10_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.SendBoleto;

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

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
            worksheet.Range["N" + TableFormat.HeaderRow].Value = "Tags";
            worksheet.Range["O" + TableFormat.HeaderRow].Value = "Descrição 1";
            worksheet.Range["P" + TableFormat.HeaderRow].Value = "Valor 1";
            worksheet.Range["Q" + TableFormat.HeaderRow].Value = "Descrição 2";
            worksheet.Range["R" + TableFormat.HeaderRow].Value = "Valor 2";
            worksheet.Range["S" + TableFormat.HeaderRow].Value = "Descrição 3";
            worksheet.Range["T" + TableFormat.HeaderRow].Value = "Valor 3";
            worksheet.Range["U" + TableFormat.HeaderRow].Value = "Nome do Sacador Avalista";
            worksheet.Range["V" + TableFormat.HeaderRow].Value = "CPF/CNPJ do Sacador Avalista";

            string returnMessage = "";
            string warningMessage = "";
            string errorMessage = "";

            int iteration = 0;
            int errorNum = 10;

            var initRow = TableFormat.HeaderRow + 1;
            lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            List<Dictionary<string, object>> boletos = new List<Dictionary<string, object>>();
            List<int> boletoNumbers = new List<int>();

            for (int row = initRow; row <= lastRow; row++)
            {
                iteration++;

                string name = worksheet.Range["A" + row].Value?.ToString();
                string taxID = worksheet.Range["B" + row].Value?.ToString();
                string streetLine1 = worksheet.Range["C" + row].Value?.ToString();
                string streetLine2 = worksheet.Range["D" + row].Value?.ToString();
                string district = worksheet.Range["E" + row].Value?.ToString();
                string city = worksheet.Range["F" + row].Value?.ToString();
                string stateCode = worksheet.Range["G" + row].Value?.ToString();
                string zipCode = worksheet.Range["H" + row].Value?.ToString();
                string amountString = worksheet.Range["I" + row].Value?.ToString();
                int amount = Convert.ToInt32(double.Parse(amountString) * 100);
                string due = worksheet.Range["J" + row].Value?.ToString();
                string fineString = worksheet.Range["K" + row].Value?.ToString();
                string interestString = worksheet.Range["L" + row].Value?.ToString().Replace(",", ".");
                string expirationString = worksheet.Range["M" + row].Value?.ToString().Replace(",", ".");
                string[] tags = worksheet.Range["N" + row].Value?.ToString().Split(',');


                List<Dictionary<string, object>> descriptions = new List<Dictionary<string, object>>();

                string description1 = worksheet.Range["O" + row].Value?.ToString();
                string value1 = worksheet.Range["P" + row].Value?.ToString();
                string description2 = worksheet.Range["Q" + row].Value?.ToString();
                string value2 = worksheet.Range["R" + row].Value?.ToString();
                string description3 = worksheet.Range["S" + row].Value?.ToString();
                string value3 = worksheet.Range["T" + row].Value?.ToString();
                string receiverName = worksheet.Range["U" + row].Value?.ToString();
                string receiverTaxId = worksheet.Range["V" + row].Value?.ToString();

                if (description1 != null && value1 != null)
                {
                    descriptions.Add(new Dictionary<string, object>
                    {
                        { "text", description1 },
                        { "amount", Convert.ToInt32(double.Parse(value1) * 100) },
                    });
                }

                if (description2 != null && value2 != null)
                {
                    descriptions.Add(new Dictionary<string, object>
                    {
                        { "text", description2 },
                        { "amount", Convert.ToInt32(double.Parse(value2) * 100) },
                    });
                }

                if (description3 != null && value3 != null)
                {
                    descriptions.Add(new Dictionary<string, object>
                    {
                        { "text", description3 },
                        { "amount", Convert.ToInt32(double.Parse(value3) * 100) },
                    });
                }

                Dictionary<string, object> boleto = new Dictionary<string, object> {
                    { "name", name },
                    { "taxId", taxID },
                    { "streetLine1", streetLine1 },
                    { "streetLine2", streetLine2 },
                    { "district", district },
                    { "city", city },
                    { "stateCode", stateCode },
                    { "zipCode", zipCode },
                    { "amount", amount },
                    { "descriptions" , descriptions },
                    { "tags", tags }
                };

                if (receiverName != null) boleto.Add("receiverName", receiverName);

                if (receiverTaxId != null) boleto.Add("receiverTaxId", receiverTaxId);

                if (due != null) boleto.Add("due", new StarkDate(due).ToString());

                if (expirationString != null) boleto.Add("overdueLimit", int.Parse(expirationString));
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
                        errorMessage = Utils.ParsingErrors(ex.Message, errorNum);
                    }

                errorNum += 100;

                nextIteration:
                    initRow = row + 1;
                    boletos = new List<Dictionary<string, object>>();
                    boletoNumbers = new List<int>();
                }
            }

            MessageBox.Show(warningMessage + returnMessage + errorMessage);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Globals.Main.Activate();
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            var worksheet = Globals.SendBoleto;

            Range range = worksheet.Range["A" + (TableFormat.HeaderRow + 1) + ":K1048576"];
            range.ClearContents();
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
