using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Forms;
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
    public partial class Planilha11
    {
        private void Planilha11_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha11_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            this.button4.Click += new System.EventHandler(this.button4_Click);
            this.button5.Click += new System.EventHandler(this.button5_Click);
            this.Startup += new System.EventHandler(this.Planilha11_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha11_Shutdown);

        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Globals.Main.Activate();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            LoginForm loginForm = new LoginForm();
            loginForm.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Utils.LogOut();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.Planilha11;

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":V" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "Cartão";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Estabelecimento";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Descrição";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Categoria";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Cardholder";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Saldo";

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";
            Dictionary<string, object> returnedData = new Dictionary<string, object>();

            do
            {
                int logRow = row;

                JObject respJson;

                try
                {
                    respJson = CorporateTransaction.Get(cursor);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                JArray transactions = (JArray)respJson["transactions"];

                string queryStr = "";
                Dictionary<string, JObject> corporateTransactionIds = new Dictionary<string, JObject>();


                foreach (JObject transaction in transactions)
                {
                    if (transaction["source"].ToString().Substring(0, 19) == "corporate-purchase/")
                    {
                        queryStr = queryStr + transaction["source"].ToString().Substring(19) + ",";
                        corporateTransactionIds.Add(transaction["id"].ToString(), transaction);
                    }
                }

                Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
                keyValuePairs.Add("ids", queryStr);

                JObject respJson2;

                try
                {
                    respJson2 = corporatePurchase.Get(cursor, keyValuePairs);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                JArray corporatePurchases = (JArray)respJson2["purchases"];

                foreach (JObject corporatePurchase in corporatePurchases)
                {
                    foreach (string purchaseIndex in corporatePurchase["corporateTransactionIds"]) 
                    {
                        try
                        {
                            JObject transaction = corporateTransactionIds[purchaseIndex];
                            JObject purchase = corporatePurchase;

                            worksheet.Range["A" + row].Value = transaction["created"].ToString();
                            worksheet.Range["B" + row].Value = purchase["cardId"].ToString();
                            worksheet.Range["C" + row].Value = purchase["merchantDisplayName"].ToString();
                            worksheet.Range["D" + row].Value = purchase["description"].ToString();
                            worksheet.Range["E" + row].Value = purchase["merchantCategoryType"].ToString();
                            worksheet.Range["F" + row].Value = purchase["holderId"].ToString();
                            worksheet.Range["G" + row].Value = transaction["amount"].ToString();
                            worksheet.Range["H" + row].Value = transaction["balance"].ToString();
                            worksheet.Range["I" + row].Value = purchase["id"].ToString();
                            worksheet.Range["J" + row].Value = transaction["id"].ToString();

                            row++;
                        }
                        catch (Exception error) {}
                    }

                }
            } while (cursor != null);
        }
    }
}
