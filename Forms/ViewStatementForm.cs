using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;

namespace StarkBankExcel.Forms
{
    public partial class ViewStatementForm : Form
    {
        public ViewStatementForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.GetStatement;

            string afterString = afterInput.Text;
            StarkDate after = new StarkDate(afterString);
            string beforeString = beforeInput.Text;
            StarkDate before = new StarkDate(beforeString);

            TimeSpan duration = (TimeSpan)(before.Value - after.Value);

            if(duration.Days > 30)
            {
                DialogResult result = MessageBox.Show("O período selecionado é superior a 30 dias. A operação pode demorar. Continuar?", 
                                                        "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No) return;
            }

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":I" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "Categoria da Transação";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Id da Categoria";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Saldo final";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Descrição";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Id da Transação";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Tarifa";
            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Tags";

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            if (afterInput.Enabled == true) optionalParam["after"] = after;
            if (beforeInput.Enabled == true) optionalParam["before"] = before;

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";

            do
            {
                JObject respJson;

                try
                {
                    respJson = Transaction.Get(cursor, optionalParam);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }

                if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                JArray transactions = (JArray)respJson["transactions"];

                foreach (JObject transaction in transactions)
                {
                    string path = transaction["source"].ToString();

                    string[] splitPath = path.Split('/');

                    List<string> tags = transaction["tags"].ToObject<List<string>>();

                    worksheet.Range["A" + row].Value = new StarkDateTime((string)transaction["created"]).ToString();
                    worksheet.Range["B" + row].Value = getTransactionType(splitPath, tags);

                    if(splitPath.Length > 1) worksheet.Range["C" + row].Value = splitPath[1];

                    Range amountCell = worksheet.Range["D" + row];
                    amountCell.Value = double.Parse((string)transaction["amount"]) / 100;

                    amountCell.Font.Color = XlRgbColor.rgbGreen;
                    if (amountCell.Value < 0) amountCell.Font.Color = XlRgbColor.rgbRed;

                    Range balanceCell = worksheet.Range["E" + row];
                    balanceCell.Value = double.Parse((string)transaction["balance"]) / 100;

                    balanceCell.Font.Color = XlRgbColor.rgbGreen;
                    if (balanceCell.Value < 0) balanceCell.Font.Color = XlRgbColor.rgbRed;

                    worksheet.Range["F" + row].Value = transaction["description"];
                    worksheet.Range["G" + row].Value = transaction["id"];
                    worksheet.Range["H" + row].Value = double.Parse((string)transaction["fee"]) / 100;
                    worksheet.Range["I" + row].Value = Utils.ListToString(transaction["tags"].ToObject<List<string>>(), ",");

                    row++;
                }

            } while (cursor != null);

            Close();
        }

        private string getTransactionType(string[] splitPath, List<string> tags)
        {
            string transactionType = "";

            switch (splitPath[0])
            {
                case "self":
                    transactionType = "Transferência interna";
                    break;
                case "charge":
                    transactionType = "Recebimento de boleto pago";
                    break;
                case "invoice":
                    transactionType = "Recebimento de cobrança Pix";
                    break;
                case "deposit":
                    transactionType = "Recebimento de depósito Pix";
                    break;
                case "charge-payment":
                    transactionType = "Pag. de boleto";
                    break;
                case "brcode-payment":
                    transactionType = "Pag. de QR Code";
                    break;
                case "bar-code-payment":
                    transactionType = "Pag. de imposto/concessionária";
                    break;
                case "utility-payment":
                    transactionType = "Pag. de concessionária com cód. de barras";
                    break;
                case "darf-payment":
                    transactionType = "Pag. de DARF sem cód. de barras";
                    break;
                case "tax-payment":
                    transactionType = "Pag. de imposto com cód. de barras";
                    break;
                case "transfer-request":
                    transactionType = "Transf. sem aprovação";
                    break;
                case "transfer":
                    
                    transactionType = "Transf. sem aprovação";

                    foreach (string tag in tags)
                    {
                        if(tag.Contains("payment-request/"))
                        {
                            transactionType = "Transf. com aprovação";
                        }
                    }
                    break;

                default:
                    transactionType = "Outros";
                    break;
            }
            
            if(IsChargeback(splitPath))
            {
                transactionType = "Estorno: " + transactionType;
            }

            return transactionType;
        }


        private bool IsChargeback(string[] splitPath)
        {
            if(splitPath.Length > 2)
            {
                if (splitPath[2] == "chargeback")
                {
                    return true;
                }
            }
            return false;
        }

        private void periodInput_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (periodInput.SelectedIndex)
            {
                case 0:
                    afterInput.Enabled = true;
                    beforeInput.Enabled = false;
                    break;
                case 1:
                    afterInput.Enabled = false;
                    beforeInput.Enabled = true;
                    break;
                case 2:
                    afterInput.Enabled = true;
                    beforeInput.Enabled = true;
                    break;
                case 3:
                    afterInput.Enabled = false;
                    beforeInput.Enabled = false;
                    break;
            }
        }

        private void ViewStatementForm_Load(object sender, EventArgs e)
        {
            periodInput.Items.Add("Data Inicial");
            periodInput.Items.Add("Data Final");
            periodInput.Items.Add("Intervalo");
            periodInput.Items.Add("Todos");

            periodInput.Text = "Intervalo";
        }
    }
}
