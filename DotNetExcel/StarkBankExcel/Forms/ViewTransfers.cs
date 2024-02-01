using System;
using System.Data;
using System.Linq;
using System.Text;
using EllipticCurve;
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
    public partial class ViewTransfers : Form
    {
        public ViewTransfers()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            var worksheet = Globals.GetTransfers;

            string afterString = afterInput.Text;
            string after = new StarkDate(afterString).ToString();
            string beforeString = beforeInput.Text;
            string before = new StarkDate(beforeString).ToString();

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":V" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data de Criação";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "Id da Transferência";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Status";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Nome";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "CPF/CNPJ";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Código do Banco";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Agência";

            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Número de Conta";
            worksheet.Range["J" + TableFormat.HeaderRow].Value = "Tipo de Conta";
            worksheet.Range["K" + TableFormat.HeaderRow].Value = "Ids de Transação (Saída, Estorno)";
            worksheet.Range["L" + TableFormat.HeaderRow].Value = "Tags";
            worksheet.Range["M" + TableFormat.HeaderRow].Value = "Descrição";
            worksheet.Range["N" + TableFormat.HeaderRow].Value = "Detalhamento de falha";

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();
            Dictionary<string, object> logsFailedByTransfer = new Dictionary<string, object>();

            string status = "";

            bool ratioChecked = false;

            if (TransactionId.Text != "")
            {
                optionalParam["transactionIds"] = TransactionId.Text.ToString();
            }

            if (TransactionId.Text == "")
            {
                if (afterInput.Enabled == true) optionalParam["after"] = after;
                if (beforeInput.Enabled == true) optionalParam["before"] = before;

                if (success.Checked) status += "success";
                if (processing.Checked) status += "processing";
                if (failed.Checked) status += "failed";
            }

            if (detail.Checked) ratioChecked = true;

            optionalParam.Add("status", status);

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";

            int logRow = 0;

            if (afterInput.Enabled == true && beforeInput.Enabled == true)
            {
                do
                {
                    logRow = row;
                    JObject respJson;
                    try
                    {
                        respJson = Transfer.Get(cursor, optionalParam);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        respJson = new JObject();
                        return;
                    }
                    
                    if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                    JArray transfers = (JArray)respJson["transfers"];

                    foreach (JObject transfer in transfers)
                    {
                        worksheet.Range["A" + row].Value = new StarkDateTime((string)transfer["created"]).Value; 
                        worksheet.Range["B" + row].Value = transfer["id"];
                        worksheet.Range["C" + row].Value = double.Parse((string)transfer["amount"]) / 100;
                        worksheet.Range["D" + row].Value = transfer["status"];
                        worksheet.Range["E" + row].Value = transfer["name"];
                        worksheet.Range["F" + row].Value = transfer["taxId"];
                        worksheet.Range["G" + row].Value = transfer["bankCode"];
                        worksheet.Range["H" + row].Value = transfer["branchCode"];
                        worksheet.Range["I" + row].Value = transfer["accountNumber"];
                        worksheet.Range["J" + row].Value = getAccountTypePT(transfer["accountType"].ToString());
                        worksheet.Range["K" + row].Value = string.Join(",", transfer["transactionIds"]);
                        worksheet.Range["L" + row].Value = string.Join(",", transfer["tags"]);
                        worksheet.Range["M" + row].Value = transfer["description"];

                        if (detail.Checked)
                        {
                            logsFailedByTransfer.Add(transfer["id"].ToString(), new Dictionary<string, object>());
                        }

                        row++;
                    }

                    if (detail.Checked)
                    {
                        Dictionary<string, object> logsParam = new Dictionary<string, object>
                        {
                            { "types", "failed" }
                        };
                        string keys = "";
                        string sep = "";

                        foreach (string transferId in logsFailedByTransfer.Keys)
                        {
                            keys = keys + sep + transferId;
                            sep = ",";
                        }

                        logsParam.Add("transferIds", keys);

                        try
                        {
                            respJson = TransferLog.Get("", logsParam);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Close();
                            return;
                        }

                        JArray transferLogs = (JArray)respJson["logs"];

                        foreach (JObject transferLog in transferLogs)
                        {
                            logsFailedByTransfer[(string)transferLog["transfer"]["id"]] = transferLog;
                        }

                        foreach (JObject transfer in transfers)
                        {
                            if ((string)transfer["status"] == "failed")
                            {
                                JObject failedTransfer = (JObject)logsFailedByTransfer[transfer["id"].ToString()];
                                worksheet.Range["N" + logRow].Value = string.Join(",", failedTransfer["errors"]);
                            }

                            logRow++;
                        }

                        logsFailedByTransfer = new Dictionary<string, object>();
                    }

                } while (cursor != null);

                Close();
                return;
            }
        }

        private string getAccountTypePT(string accountType)
        {

            string getAccountNumber = "";

            if (accountType == "checking") getAccountNumber = "corrente";

            if (accountType == "savings") getAccountNumber = "poupança";

            if (accountType == "payment") getAccountNumber = "pagamento";

            if (accountType == "salary") getAccountNumber = "salario";

            return getAccountNumber;
        }
    }
}
