using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace StarkBankExcel.Forms
{
    public partial class ViewInvoiceForm : Form
    {
        public ViewInvoiceForm()
        {
            InitializeComponent();
        }

        private void ViewInvoiceForm_Load(object sender, EventArgs e)
        {
            statusInput.Items.Add("Todos");
            statusInput.Items.Add("Pagos");
            statusInput.Items.Add("Criados");
            statusInput.Items.Add("Vencidos");
            statusInput.Items.Add("Cancelados");
            statusInput.Items.Add("Expirados");

            statusInput.Text = "Todos";

            periodInput.Text = "Intervalo";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.GetInvoices;

            string afterString = afterInput.Text;
            string after = new StarkDate(afterString).ToString();
            string beforeString = beforeInput.Text;
            string before = new StarkDate(beforeString).ToString();

            string statusString = statusInput.Text;

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data de Emissão";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "Nome";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "CPF/CNPJ";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Status";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Valor de Emissão";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Desconto";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Multa";
            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Juros";
            worksheet.Range["J" + TableFormat.HeaderRow].Value = "Vencimento";
            worksheet.Range["K" + TableFormat.HeaderRow].Value = "Pagável até";
            worksheet.Range["L" + TableFormat.HeaderRow].Value = "Copia e Cola (BR Code)";
            worksheet.Range["M" + TableFormat.HeaderRow].Value = "Id da Invoice";
            worksheet.Range["N" + TableFormat.HeaderRow].Value = "Tarifa";
            worksheet.Range["O" + TableFormat.HeaderRow].Value = "Tags";
            worksheet.Range["P" + TableFormat.HeaderRow].Value = "Link PDF";

            string status = GetStatus(statusString);

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            if (status != "all") optionalParam["status"] = status;

            if (afterInput.Enabled == true) optionalParam["after"] = after;
            if (beforeInput.Enabled == true) optionalParam["before"] = before; 

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";

            do
            {
                JObject respJson;

                try
                {
                    respJson = Invoice.Get(cursor, optionalParam);
                } catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }

                if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                JArray invoices = (JArray)respJson["invoices"];

                foreach (JObject invoice in invoices)
                {
                    worksheet.Range["A" + row].Value = new StarkDateTime((string)invoice["created"]).ToString();
                    worksheet.Range["B" + row].Value = invoice["name"];
                    worksheet.Range["C" + row].Value = invoice["taxId"];
                    worksheet.Range["D" + row].Value = GetStatusInPt((string)invoice["status"]);
                    worksheet.Range["E" + row].Value = double.Parse((string)invoice["amount"]) / 100;
                    worksheet.Range["F" + row].Value = double.Parse((string)invoice["nominalAmount"]) / 100;
                    worksheet.Range["G" + row].Value = double.Parse((string)invoice["discountAmount"]) / 100;
                    worksheet.Range["H" + row].Value = double.Parse((string)invoice["fineAmount"]) / 100;
                    worksheet.Range["I" + row].Value = double.Parse((string)invoice["interestAmount"]) / 100;
                    worksheet.Range["J" + row].Value = invoice["due"];
                    worksheet.Range["K" + row].Value = invoice["expiration"];
                    worksheet.Range["L" + row].Value = invoice["brcode"];
                    worksheet.Range["M" + row].Value = invoice["id"];
                    worksheet.Range["N" + row].Value = double.Parse((string)invoice["fee"]) / 100;
                    worksheet.Range["O" + row].Value = Utils.ListToString(invoice["tags"].ToObject<List<string>>(), ",");

                    Microsoft.Office.Interop.Excel.Range rng = worksheet.Range["P" + row];
                    rng.Value = "PDF";
                    Microsoft.Office.Interop.Excel.Hyperlink link = rng.Hyperlinks.Add(rng, (string)invoice["pdf"]);

                    row++;
                }

            } while (cursor != null);

            Close();
        }

        private string GetStatus(string status)
        {
            switch (status)
            {
                case "Pagos":
                    return "paid";
                case "Criados":
                    return "created";
                case "Vencidos":
                    return "overdue";
                case "Cancelados":
                    return "canceled";
                case "Expirados":
                    return "expired";
                default:
                    return "all";
            }
        }

        private string GetStatusInPt(string status)
        {
            switch (status)
            {
                case "paid":
                    return "pago";
                case "voided":
                    return "anulado";
                case "created":
                    return "criado";
                case "overdue":
                    return "vencido";
                case "canceled":
                    return "cancelado";
                case "expired":
                    return "expirado";
                case "unknown":
                    return "desconhecido";
                default:
                    return "status inválido";
            }
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
    }
}
