using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
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
using System.IO;
using System.Configuration;
using System.Linq;
using System.Net.Mail;

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
            this.button6.Click += new System.EventHandler(this.button6_Click);

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
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "ID Compra";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Categoria";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Estabelecimento";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Descrição Compra";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Status";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Anexo";
            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Descrição Anexo";
            worksheet.Range["J" + TableFormat.HeaderRow].Value = "ID Cartão";
            worksheet.Range["K" + TableFormat.HeaderRow].Value = "ID Holder";
            worksheet.Range["L" + TableFormat.HeaderRow].Value = "ID Centro de Custo";

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";
            Dictionary<string, object> returnedData = new Dictionary<string, object>();

            do
            {

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
                    worksheet.Range["B" + row].Value = purchase["id"];
                    worksheet.Range["C" + row].Value = purchase["merchantCategoryCode"];
                    worksheet.Range["D" + row].Value = purchase["merchantName"];
                    worksheet.Range["E" + row].Value = purchase["description"];
                    worksheet.Range["F" + row].Value = purchase["status"];
                    worksheet.Range["G" + row].Value = purchase["amount"];

                    try
                    {
                        foreach (JToken attachment in purchase["attachments"])
                        {
                            worksheet.Range["H" + row].Value = attachment["id"] + "," + worksheet.Range["H" + row].Value;
                        }
                        worksheet.Range["H" + row].Value = worksheet.Range["H" + row].Value.Substring(0, worksheet.Range["H" + row].Value.Length - 1);
                    }
                    catch { }

                    worksheet.Range["I" + row].Value = "descrição anexo";
                    worksheet.Range["J" + row].Value = purchase["cardId"];
                    worksheet.Range["K" + row].Value = purchase["holderId"];
                    worksheet.Range["L" + row].Value = purchase["cenderId"];

                    row++;
                }

            } while (cursor != null);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.Planilha12;

            Excel.Range selectedRange = worksheet.Application.Selection;

            worksheet.Range["H" + 6].Value = selectedRange.Address;

            if (selectedRange.Address.Substring(1, 1) == "H")
            {
                JObject respJson;
                int range_a = int.Parse(selectedRange.Address.Substring(3, 2));
                int range_b = int.Parse(selectedRange.Address.Substring(9, 2));
                bool validator = false;


                for (int i = range_a; (i - 1) < range_b; i += 1)
                {

                    if (worksheet.Range["H" + i].Value != null)
                    {
                        string[] attachmentIds = worksheet.Range["H" + i].Value.Split(',');

                        int fileNumber = 0;

                        foreach (string attachmentId in attachmentIds)
                        {
                            try
                            {
                                respJson = corporateAttachment.Get(attachmentId);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            JObject attachmentFile = (JObject)respJson["attachment"];
                            string attachmentString = attachmentFile["content"].ToString();
                            string attachment = attachmentString.Substring(attachmentString.IndexOf("base64,") + "base64,".Length);
                            string[] parts = attachmentString.Split(new[] { ";base64," }, StringSplitOptions.None);

                            if (parts.Length == 2)
                            {
                                validator = true;

                                string contentType = parts[0].Split(':')[1];
                                string extension = contentType.Split('/')[1];

                                worksheet.Range["H" + 7].Value = attachment;

                                byte[] attachmentb64 = Convert.FromBase64String(attachment);

                                string path = Environment.CurrentDirectory + "\\Anexos\\";
                                string fileName = worksheet.Range["A" + i].Value.Substring(0, 10).Replace("-", "") + "-" + worksheet.Range["B" + i].Value + "-" + worksheet.Range["D" + i].Value;

                                if (!Directory.Exists(path))
                                {
                                    Directory.CreateDirectory(path);
                                }

                                if (fileNumber == 0)
                                {
                                    File.WriteAllBytes(path + fileName + "." + extension, attachmentb64);
                                }
                                else
                                {
                                    File.WriteAllBytes(path + fileName + $" ({fileNumber})" + "." + extension, attachmentb64);
                                }
                            }

                            fileNumber++;
                        }
                    }
                }

                if (validator == true)
                {
                    MessageBox.Show("Os Anexos selecionados foram salvos");
                }

            }
        }
    }
}
