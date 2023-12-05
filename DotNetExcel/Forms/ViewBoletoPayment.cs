using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using System.Threading.Tasks;
using StarkBankExcel.Resources;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace StarkBankExcel.Forms
{
    public partial class ViewBoletoPayment : Form
    {
        public ViewBoletoPayment()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.GetBoletoPayment;

            string afterString = afterInput.Text;
            string after = new StarkDate(afterString).ToString();
            string beforeString = beforeInput.Text;
            string before = new StarkDate(beforeString).ToString();

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":V" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data de Criação";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Status";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Data de Agendamento";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Linha Digitável";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Descrição";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Tags";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Id do Pagamento";

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            if (afterInput.Enabled == true) optionalParam["after"] = after;
            if (beforeInput.Enabled == true) optionalParam["before"] = before;

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";

            Dictionary<string, object> returnedData = new Dictionary<string, object>();

            if (afterInput.Enabled == true && beforeInput.Enabled == true)
            {
                do
                {
                    JObject respJson;

                    try
                    {
                        respJson = BoletoPaymentClass.Get(cursor, optionalParam);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                    JArray payments = (JArray)respJson["payments"];

                    foreach (JObject payment in payments)
                    {
                        worksheet.Range["A" + row].Value = payment["created"];
                        worksheet.Range["B" + row].Value = double.Parse((string)payment["amount"]) / 100;
                        worksheet.Range["C" + row].Value = payment["status"];
                        worksheet.Range["D" + row].Value = payment["scheduled"];
                        worksheet.Range["E" + row].Value = payment["line"];
                        worksheet.Range["F" + row].Value = payment["description"];
                        worksheet.Range["G" + row].Value = Utils.ListToString(payment["tags"].ToObject<List<string>>(), ",");
                        worksheet.Range["H" + row].Value = payment["id"];

                        row++;
                    }
                } while (cursor != null);

                Close();
                return;
            }
        }

        private void periodInput_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void afterInput_ValueChanged(object sender, EventArgs e)
        {

        }

        private void beforeInput_ValueChanged(object sender, EventArgs e)
        {

        }

    }
}
