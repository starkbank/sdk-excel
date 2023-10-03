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
    public partial class Planilha12
    {
        private void Planilha12_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha12_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Globals.Main.Activate();
        }

        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            this.button3.Click += new System.EventHandler(this.button3_Click);
            this.button4.Click += new System.EventHandler(this.button4_Click);
            this.button5.Click += new System.EventHandler(this.button5_Click);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            LoginForm loginForm = new LoginForm();
            loginForm.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Utils.LogOut();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            var worksheet = Globals.Planilha12;

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
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Status";

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
                    respJson = corporatePurchase.Get(cursor);
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
                    worksheet.Range["B" + row].Value = purchase["cardId"];
                    worksheet.Range["C" + row].Value = purchase["merchantDisplayName"];
                    worksheet.Range["D" + row].Value = purchase["description"];
                    worksheet.Range["E" + row].Value = purchase["merchantCategoryType"];
                    worksheet.Range["F" + row].Value = purchase["holderId"];
                    worksheet.Range["G" + row].Value = purchase["amount"];
                    worksheet.Range["H" + row].Value = purchase["status"];

                    row++;
                }

            } while (cursor != null);
        }
    }
}
