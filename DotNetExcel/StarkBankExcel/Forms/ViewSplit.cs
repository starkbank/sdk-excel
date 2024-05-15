using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StarkBankExcel.Forms
{
    public partial class ViewSplit : Form
    {
        public ViewSplit()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string afterString = afterInput.Text;
            string after = new StarkDate(afterString).ToString();
            string beforeString = beforeInput.Text;
            string before = new StarkDate(beforeString).ToString();
            string statusString = statusInput.Text;

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

            string status = GetStatus(statusString);

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            if (status != "all") optionalParam["status"] = status;

            if (afterInput.Enabled == true) optionalParam["after"] = after;
            if (beforeInput.Enabled == true) optionalParam["before"] = before;

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";
            Dictionary<string, object> returnedData = new Dictionary<string, object>();

            do
            {
                Dictionary<string, string> splitsById = new Dictionary<string, string>();
                List<string> receiverList = new List<string>();
                int logRow = row;

                JObject respJson;

                try
                {
                    respJson = Split.Get(cursor, optionalParam);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                JArray splitsObjs = (JArray)respJson["splits"];

                foreach (JObject split in splitsObjs)
                {
                    receiverList.Add(split["receiverId"].ToString());
                }

                try
                {
                    string receiversString = string.Join(",", receiverList);
                    respJson = SplitReceiver.Get("", new Dictionary<string, object>() { { "ids", receiversString } });

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                foreach (JObject receiver in (JArray)respJson["receivers"]) 
                {
                    splitsById.Add(receiver["id"].ToString(), receiver["taxId"].ToString());
                }

                List<string> checkReceiverList = new List<string>();
                Dictionary<string, Dictionary<string, int>> directionIds = new Dictionary<string, Dictionary<string, int>>();
                int column = 7;

                foreach (JObject receiver in splitsObjs)
                {
                    worksheet.Range["A" + row].Value = receiver["created"].ToString();
                    worksheet.Range["B" + row].Value = receiver["source"].ToString().Substring(8);
                    worksheet.Range["C" + row].Value = receiver["receiverId"].ToString();
                    worksheet.Range["D" + row].Value = splitsById[receiver["receiverId"].ToString()].ToString();
                    worksheet.Range["E" + row].Value = double.Parse((string)receiver["amount"].ToString()) / 100;
                    worksheet.Range["F" + row].Value = receiver["status"].ToString();

                    row++;
                }

            } while (cursor != null);
        }

        private string GetStatus(string status)
        {
            switch (status)
            {
                case "Criado":
                    return "created";
                case "Processando":
                    return "processing";
                case "Cancelado":
                    return "canceled";
                case "Sucesso":
                    return "success";
                case "Falhas":
                    return "failed";
                default:
                    return "all";
            }
        }
    }
}
