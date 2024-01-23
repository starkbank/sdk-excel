using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using StarkBankExcel.Forms;
using StarkBankExcel.Resources;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace StarkBankExcel
{
    public partial class GetBoleto
    {
        private void Planilha8_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha8_Shutdown(object sender, System.EventArgs e)
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
            this.button6.Click += new System.EventHandler(this.button6_Click);
            this.button7.Click += new System.EventHandler(this.button7_Click);
            this.Startup += new System.EventHandler(this.Planilha8_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha8_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            ViewBoletoForm viewChargeForm = new ViewBoletoForm();
            viewChargeForm.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Globals.Main.Activate();
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

        private void button6_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.GetBoleto;


            Range range = worksheet.Range["A" + (TableFormat.HeaderRow + 1) + ":K1048576"];
            range.ClearContents();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.GetBoleto;
            string selectedPath = "";

            Excel.Range selectedRange = worksheet.Application.Selection;

            string cell_A = selectedRange.Address.Split('$')[1];
            string cell_J = selectedRange.Address.Split('$')[3];

            if (cell_A != "A" | cell_J != "V")
            {
                MessageBox.Show("Todas as colunas devem ser selecionadas !");
            }

            if (cell_A == "A" & cell_J == "V")
            {
                byte[] respJson;
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
                    if (worksheet.Range["M" + i].Value != null)
                    {
                        validator = false;
                        string boletoId = worksheet.Range["M" + i].Value;
                        string created = new StarkDateTime(worksheet.Range["A" + i].Value).Value.ToString();
                        string value = worksheet.Range["E" + i].Value;
                        string name = worksheet.Range["B" + i].Value;

                        try
                        {
                            respJson = Boleto.Pdf(boletoId);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        string fileName = "boleto-" + boletoId + ".pdf";

                        using (FileStream fs = File.Create(selectedPath + "\\" + fileName))
                        {
                            fs.Write(respJson, 0, respJson.Length);
                            validator = true;
                        }
                    }
                }

                if (validator == true)
                {
                    MessageBox.Show("Arquivos salvos em: " + selectedPath);
                }

            }
        }
    }
}
