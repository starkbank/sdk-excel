using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using System.Threading.Tasks;
using StarkBankExcel.Resources;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace StarkBankExcel.Forms
{
    public partial class cardStatmentForm : Form
    {
        public cardStatmentForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string afterString = afterInput.Text;
            string after = new StarkDate(afterString).ToString();
            string beforeString = beforeInput.Text;
            string before = new StarkDate(beforeString).ToString();

            var worksheet = Globals.Planilha11;

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":V" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "ID";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Estabelecimento";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Descrição";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Saldo";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Source";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Tags";

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            optionalParam.Add("after", after);
            optionalParam.Add("before", before);

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";
            Dictionary<string, object> returnedData = new Dictionary<string, object>();

            do
            {
                Dictionary<string, JObject> corporatePurchaseById = new Dictionary<string, JObject>();
                int logRow = row;

                JObject respJson;

                try
                {
                    respJson = CorporateTransaction.Get(cursor, optionalParam);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                JArray transactions = (JArray)respJson["transactions"];

                string query = "";

                foreach (JObject transaction in transactions)
                {
                    if (transaction["source"].ToString().Substring(0, 19) == "corporate-purchase/")
                    {
                        query = query + transaction["source"].ToString().Substring(19) + ",";
                    }
                }

                Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
                keyValuePairs.Add("ids", query);

                JObject respJson2;

                try
                {
                    respJson2 = corporatePurchase.Get(null, keyValuePairs);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                JArray corporatePurchases = (JArray)respJson2["purchases"];

                foreach (JObject corporatePurchase in corporatePurchases)
                {
                    corporatePurchaseById.Add(corporatePurchase["id"].ToString(), corporatePurchase);
                }

                foreach (JObject transaction in transactions)
                {
                    if (transaction["source"].ToString().Substring(0, 19) == "corporate-purchase/")
                    {
                        string transactionIndx = transaction["source"].ToString().Substring(19);
                        JObject purchase = corporatePurchaseById[transactionIndx];

                        worksheet.Range["A" + row].Value = transaction["created"].ToString();
                        worksheet.Range["B" + row].Value = transaction["id"].ToString();
                        worksheet.Range["C" + row].Value = purchase["merchantName"].ToString();
                        worksheet.Range["D" + row].Value = transaction["description"].ToString();
                        worksheet.Range["E" + row].Value = (double.Parse((string)transaction["amount"]) / 100).ToString().ToString();
                        worksheet.Range["F" + row].Value = (double.Parse((string)transaction["balance"]) / 100).ToString();
                        worksheet.Range["G" + row].Value = transaction["source"].ToString();
                        worksheet.Range["H" + row].Value = Utils.ListToString(transaction["tags"].ToObject<List<string>>(), ",");

                        row++;
                    }
                }
            } while (cursor != null);

            Close();
            return;
        }
    }
}
