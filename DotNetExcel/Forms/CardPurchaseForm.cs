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
    public partial class CardPurchaseForm : Form
    {
        public CardPurchaseForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string afterString = afterInput.Text;
            string after = new StarkDate(afterString).ToString();
            string beforeString = beforeInput.Text;
            string before = new StarkDate(beforeString).ToString();

            var worksheet = Globals.Planilha12;

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":V" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "ID Compra";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Categoria";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Estabelecimento";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Descrição Compra";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Status";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Anexo";
            worksheet.Range["I" + TableFormat.HeaderRow].Value = "ID Cartão";
            worksheet.Range["J" + TableFormat.HeaderRow].Value = "ID Holder";

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";
            Dictionary<string, object> returnedData = new Dictionary<string, object>();

            string status = "";

            if (OptionButtonConfirmed.Checked) optionalParam.Add("status", "confirmed");
            if (OptionButtonApproved.Checked) optionalParam.Add("status", "approved");
            if (OptionButtonCanceled.Checked) optionalParam.Add("status", "canceled");
            if (Denied.Checked) optionalParam.Add("status", "denied");
            if (Voided.Checked) optionalParam.Add("status", "voided");

            optionalParam.Add("after", after);
            optionalParam.Add("before", before);

            do
            {
                JObject respJson;

                try
                {
                    respJson = corporatePurchase.Get(cursor, optionalParam);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                JArray purchases = (JArray)respJson["purchases"];

                foreach (JObject purchase in purchases)
                {
                    worksheet.Range["A" + row].Value = purchase["created"];
                    worksheet.Range["B" + row].Value = purchase["id"];
                    worksheet.Range["C" + row].Value = purchase["merchantCategoryCode"];
                    worksheet.Range["D" + row].Value = purchase["merchantName"];
                    worksheet.Range["E" + row].Value = purchase["description"];
                    worksheet.Range["F" + row].Value = purchase["status"];
                    worksheet.Range["G" + row].Value = double.Parse((string)purchase["amount"]) / 100;

                    foreach (JToken attachment in purchase["attachments"])
                    {
                        worksheet.Range["H" + row].Value = attachment["id"] + "," + worksheet.Range["H" + row].Value;
                    }

                    if (worksheet.Range["H" + row].Value != null)
                    {
                        worksheet.Range["H" + row].Value = worksheet.Range["H" + row].Value.Substring(0, worksheet.Range["H" + row].Value.Length - 1);
                    }

                    worksheet.Range["I" + row].Value = purchase["cardId"];
                    worksheet.Range["J" + row].Value = purchase["holderId"];
                    worksheet.Range["K" + row].Value = purchase["cenderId"];

                    row++;
                }

            } while (cursor != null);

            Close();
            return;
        }
    }
}
