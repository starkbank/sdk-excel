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
            var worksheet = Globals.GetSplit;

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":V" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "ID da Invoice";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Seller";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "CPF/CNPJ";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Status";

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";
            Dictionary<string, object> returnedData = new Dictionary<string, object>();

            do
            {
                Dictionary<string, string> receiversById = new Dictionary<string, string>();
                List<string> receiversList = new List<string>();
                int logRow = row;

                JObject respJson;

                try
                {
                    respJson = SplitReceiver.Get(cursor);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                JArray receivers = (JArray)respJson["receivers"];

                foreach (JObject receiver in receivers)
                {
                    receiversById.Add(receiver["id"].ToString(), receiver["taxId"].ToString());
                    receiversList.Add(receiver["id"].ToString());
                }

                try
                {
                    string receiversString = string.Join(",", receiversList);
                    respJson = Split.Get("", new Dictionary<string, object>() { { "receiverIds", receiversString } });

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                receivers = (JArray)respJson["splits"];

                // vai ter que implemetar uma caralhada de condição e codigo aqui pra fazer essa parada que foi alinhada por produto

                List<string> checkReceiverList = new List<string>();
                Dictionary<string, Dictionary<string, int>> directionIds = new Dictionary<string, Dictionary<string, int>>();
                int column = 7;

                foreach (JObject receiver in receivers)
                {
                    worksheet.Range["A" + row].Value = receiver["created"].ToString();
                    worksheet.Range["B" + row].Value = receiver["source"].ToString().Substring(8);
                    worksheet.Range["C" + row].Value = receiver["receiverId"].ToString();
                    worksheet.Range["D" + row].Value = receiversById[receiver["receiverId"].ToString()].ToString();
                    worksheet.Range["E" + row].Value = double.Parse((string)receiver["amount"].ToString()) / 100;
                    worksheet.Range["F" + row].Value = receiver["status"].ToString();

                    row++;
                }

            } while (cursor != null);
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
