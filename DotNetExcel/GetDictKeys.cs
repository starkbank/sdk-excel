using System;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

namespace StarkBankExcel
{
    public partial class GetDictKeys
    {
        private void Planilha3_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha3_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            this.button2.Click += new System.EventHandler(this.button2_Click);
            this.button3.Click += new System.EventHandler(this.button3_Click);
            this.button4.Click += new System.EventHandler(this.button4_Click);
            this.button5.Click += new System.EventHandler(this.button5_Click);
            this.button6.Click += new System.EventHandler(this.button6_Click);
            this.Startup += new System.EventHandler(this.Planilha3_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha3_Shutdown);

        }

        #endregion

        private void button1_Click_1(object sender, EventArgs e)
        {
            var worksheet = Globals.GetDictKeys;

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Chave Pix";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Tags";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Descrição";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Nome";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "CPF/CNPJ";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "ISPB";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Agência";
            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Conta";
            worksheet.Range["J" + TableFormat.HeaderRow].Value = "Tipo de Conta";
            worksheet.Range["K" + TableFormat.HeaderRow].Value = "externalId";

            int initRow = TableFormat.HeaderRow + 1;
            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            Parallel.For(initRow, lastRow + 1, rowIndex =>
            {

                string keyId = worksheet.Range["A" + rowIndex].Value;
                
                List<Dictionary<string, object>> jsonData = new List<Dictionary<string, object>>();

                JObject resp;

                try
                {
                    DateTime currentTime = DateTime.UtcNow;
                    jsonData.Add(
                        new Dictionary<string, object>
                        {
                            {"id", rowIndex },
                            {"keyId", keyId },
                            {"Time", currentTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") }
                        }
                    );
                    resp = DictKey.Get(keyId);

                    JObject dictKey = (JObject)resp["key"];
                    worksheet.Range["E" + rowIndex].Value = dictKey["name"];
                    worksheet.Range["F" + rowIndex].Value = dictKey["taxId"];
                    worksheet.Range["G" + rowIndex].Value = dictKey["ispb"];
                    worksheet.Range["H" + rowIndex].Value = dictKey["branchCode"];
                    worksheet.Range["I" + rowIndex].Value = dictKey["accountNumber"];
                    worksheet.Range["J" + rowIndex].Value = dictKey["accountType"];
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), "Erro na Requisiçâo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string jsonString = Json.Encode(jsonData);
            });

            MoveToTransfer();
        }

        public static void MoveToTransfer()
        {
            var worksheet = Globals.GetDictKeys;

            var initRow = TableFormat.HeaderRow + 1;
            var lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            List<Dictionary<string, object>> validTransfers = new List<Dictionary<string, object>>();

            for (int row = initRow; row <= lastRow; row++)
            {
                string valueString = worksheet.Range["B" + row].Value?.ToString();
                int? value = null;
                if(valueString != null)
                {
                    value = int.Parse(valueString);
                }
                string tags = worksheet.Range["C" + row].Value?.ToString();
                string description = worksheet.Range["D" + row].Value?.ToString();

                string name = worksheet.Range["E" + row].Value?.ToString();
                string taxID = worksheet.Range["F" + row].Value?.ToString();
                string ispb = worksheet.Range["G" + row].Value?.ToString();
                string branchCode = worksheet.Range["H" + row].Value?.ToString();
                string accountNumber = worksheet.Range["I" + row].Value?.ToString();
                string accountType = worksheet.Range["J" + row].Value?.ToString();

                if(name != null && accountType != null)
                {
                    validTransfers.Add(new Dictionary<string, object> {
                            {"Nome", name },
                            {"CPF/CNPJ", taxID },
                            {"Valor", value },
                            {"ISPB", ispb },
                            {"Agência", branchCode },
                            {"Conta", accountNumber },
                            {"Tipo de Conta", accountType },
                            {"Tags", tags },
                            {"Descrição", description },
                        }
                    );
                }
            }

            if(validTransfers.Count == 0)
            {
                MessageBox.Show("Não há nenhuma Chave Pix válida para mover para a aba de Transferências com Aprovação", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            DialogResult result = MessageBox.Show("Foram encontradas " + validTransfers.Count + " Chaves Pix válidas. " +
                "Deseja mover para a aba de Transferências com Aprovação? Dados na aba de Transferências com Aprovação serão apagados.",
                "Deseja prosseguir?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (result == DialogResult.No)
            {
                return;
            }

            var transferWorksheet = Globals.Transfers;
            transferWorksheet.Activate();

            transferWorksheet.Range["A" + (TableFormat.HeaderRow + 1) + ":I" + transferWorksheet.Rows.Count].ClearContents();

            int transferRow = TableFormat.HeaderRow + 1;
            foreach(Dictionary<string, object> transfer in validTransfers)
            {
                transferWorksheet.Range["A" + transferRow].Value = transfer["Nome"];
                transferWorksheet.Range["B" + transferRow].Value = transfer["CPF/CNPJ"];
                transferWorksheet.Range["C" + transferRow].Value = transfer["Valor"];
                transferWorksheet.Range["D" + transferRow].Value = transfer["ISPB"];
                transferWorksheet.Range["E" + transferRow].Value = transfer["Agência"];
                transferWorksheet.Range["F" + transferRow].Value = transfer["Conta"];
                transferWorksheet.Range["G" + transferRow].Value = transfer["Tipo de Conta"];
                transferWorksheet.Range["H" + transferRow].Value = transfer["Tags"];
                transferWorksheet.Range["I" + transferRow].Value = transfer["Descrição"];
                transferRow++;
            }

            MessageBox.Show("Transferências movidas para a aba de Transferências com Aprovação", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MoveToTransfer();
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
            var worksheet = Globals.GetDictKeys;

            Range range = worksheet.Range["A" + (TableFormat.HeaderRow + 1) + ":K1048576"];
            range.ClearContents();
        }
    }
}
