using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Drawing;
using System.Net.Mail;
using System.Reflection;
using System.Diagnostics;
using System.Configuration;
using StarkBankExcel.Forms;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System.Collections.Generic;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace StarkBankExcel
{
    public partial class Planilha12
    {
        private void button4_Click(object sender, EventArgs e)
        {
            Globals.Main.Activate();
        }

        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
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

        private void button1_Click(object sender, EventArgs e)
        {
            CardPurchaseForm cardPurchase = new CardPurchaseForm();
            cardPurchase.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.Planilha12;
            string selectedPath = "";

            Excel.Range selectedRange = worksheet.Application.Selection;

            string cell_A = selectedRange.Address.Split('$')[1];
            string cell_J = selectedRange.Address.Split('$')[3];

            if (cell_A != "A" | cell_J != "J")
            {
                MessageBox.Show("Todas as colunas devem ser selecionadas !");
            }

            if (cell_A == "A" & cell_J == "J")
            {
                JObject respJson;
                int start_range = int.Parse(selectedRange.Address.Substring(3).Split(':')[0]);
                int end_range = int.Parse(selectedRange.Address.Split('$')[4]);
                bool validator = false;

                using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
                {
                    folderBrowserDialog.Description = "Selecione uma pasta";
                    folderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer;

                    if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                    {
                        selectedPath = folderBrowserDialog.SelectedPath;
                    }
                }

                for (int i = start_range; (i - 1) < end_range; i += 1)
                {
                    if (worksheet.Range["H" + i].Value != null)
                    {
                        string[] attachmentIds = worksheet.Range["H" + i].Value.Split(',');

                        int fileNumber = 0;

                        foreach (string attachmentId in attachmentIds)
                        {
                            try
                            {
                                respJson = corporateAttachment.Get(attachmentId, optionalParams: new Dictionary<string, object> { { "expand", "content" } });
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

                                byte[] attachmentb64 = Convert.FromBase64String(attachment);

                                string fileName = worksheet.Range["A" + i].Value.Substring(0, 10).Replace("-", "") + "-" + worksheet.Range["B" + i].Value + "-" + worksheet.Range["D" + i].Value;
                                fileName = Regex.Replace(fileName, "[*|@|*|&]", string.Empty);

                                if (fileNumber == 0)
                                {
                                    File.WriteAllBytes(selectedPath + "\\" + fileName + "." + extension, attachmentb64);
                                }
                                else
                                {
                                    File.WriteAllBytes(selectedPath + fileName + $" ({fileNumber})" + "." + extension, attachmentb64);
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
