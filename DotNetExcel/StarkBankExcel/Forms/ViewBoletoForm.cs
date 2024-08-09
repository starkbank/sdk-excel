using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StarkBankExcel.Forms
{
    public partial class ViewBoletoForm : Form
    {
        public ViewBoletoForm()
        {
            InitializeComponent();
        }

        private void ViewChargeForm_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            periodInput.Text = "Intervalo";
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            var worksheet = Globals.GetBoleto;

            Dictionary<string, object> logsPaidByCharge = new Dictionary<string, object>();
            Dictionary<string, object> logsRegisteredByCharge = new Dictionary<string, object>();

            string afterString = afterInput.Text;
            string after = new StarkDate(afterString).ToString();
            string beforeString = beforeInput.Text;
            string before = new StarkDate(beforeString).ToString();

            string statusString = comboBox1.Text;

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":V" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data de Emissão";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "Nome";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "CPF/CNPJ";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Status";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Valor de Emissão";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Desconto";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Multa";
            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Juros";
            worksheet.Range["J" + TableFormat.HeaderRow].Value = "Data de Crédito";
            worksheet.Range["K" + TableFormat.HeaderRow].Value = "Vencimento";
            worksheet.Range["L" + TableFormat.HeaderRow].Value = "Linha Digitável";
            worksheet.Range["M" + TableFormat.HeaderRow].Value = "Id do Boleto";
            worksheet.Range["N" + TableFormat.HeaderRow].Value = "Tarifa";
            worksheet.Range["O" + TableFormat.HeaderRow].Value = "Tags";
            worksheet.Range["P" + TableFormat.HeaderRow].Value = "Link PDF";
            worksheet.Range["Q" + TableFormat.HeaderRow].Value = "Logradouro";
            worksheet.Range["R" + TableFormat.HeaderRow].Value = "Complemento";
            worksheet.Range["S" + TableFormat.HeaderRow].Value = "Bairro";
            worksheet.Range["T" + TableFormat.HeaderRow].Value = "Cidade";
            worksheet.Range["U" + TableFormat.HeaderRow].Value = "Estado";
            worksheet.Range["V" + TableFormat.HeaderRow].Value = "CEP";

            string status = GetStatus(statusString);

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            if (status != "all") optionalParam["status"] = status;

            if (afterInput.Enabled == true) optionalParam["after"] = after;
            if (beforeInput.Enabled == true) optionalParam["before"] = before;

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";

            do
            {
                int logRow = row;

                JObject respJson;

                try
                {
                    respJson = Boleto.Get(cursor, optionalParam);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }

                if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                JArray boletos = (JArray)respJson["boletos"];

                foreach (JObject boleto in boletos)
                {
                    string boletoStatus = (string)boleto["status"];
                    string id = (string)boleto["id"];

                    worksheet.Range["A" + row].Value = new StarkDateTime((string)boleto["created"]).Value.ToString();
                    worksheet.Range["B" + row].Value = boleto["name"];
                    worksheet.Range["C" + row].Value = boleto["taxId"];
                    worksheet.Range["D" + row].Value = GetStatusInPt(boletoStatus);
                    worksheet.Range["E" + row].Value = double.Parse((string)boleto["amount"]) / 100;

                    worksheet.Range["K" + row].Value = new StarkDateTime((string)boleto["due"]).Value.ToString();
                    worksheet.Range["L" + row].Value = boleto["line"];
                    worksheet.Range["M" + row].Value = id;
                    worksheet.Range["N" + row].Value = double.Parse((string)boleto["fee"]) / 100;
                    worksheet.Range["O" + row].Value = Utils.ListToString(boleto["tags"].ToObject<List<string>>(), ",");

                    Range rng = worksheet.Range["P" + row];
                    rng.Value = "PDF";
                    Hyperlink link = rng.Hyperlinks.Add(rng, Utils.BaseUrl(Globals.Credentials.Range["B3"].Value) + "/v2/boleto/" + id + "/pdf");

                    if (OptionButtonEventCredited.Checked)
                    {
                        worksheet.Range["Q" + row].Value = boleto["streetLine1"];
                        worksheet.Range["R" + row].Value = boleto["streetLine2"];
                        worksheet.Range["S" + row].Value = boleto["district"];
                        worksheet.Range["T" + row].Value = boleto["city"];
                        worksheet.Range["U" + row].Value = boleto["stateCode"];
                        worksheet.Range["V" + row].Value = boleto["zipCode"];

                        if (boletoStatus == "paid")
                        {
                            logsPaidByCharge.Add(id, new Dictionary<string, object>());
                            logsRegisteredByCharge.Add(id, new Dictionary<string, object>());
                        }
                    }
                    row++;
                }

                if (OptionButtonEventCredited.Checked)
                {
                    Dictionary<string, object> logsParam = new Dictionary<string, object>
                    {
                        { "types", "paid" }
                    };
                    string keys = "";
                    string sep = "";

                    foreach (string boletoId in logsPaidByCharge.Keys)
                    {
                        keys = keys + sep + boletoId;
                        sep = ",";
                    }

                    logsParam.Add("boletoIds", keys);

                    string logsCursor = "";

                    JArray boletoLogs;

                    do
                    {
                        try
                        {
                            respJson = Boleto.Log.Get(logsCursor, logsParam);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Close();
                            return;
                        }

                        boletoLogs = (JArray)respJson["logs"];
                        if ((string)respJson["cursor"] != "") logsCursor = (string)respJson["cursor"];

                        foreach (JObject boletoLog in boletoLogs)
                        {
                            logsPaidByCharge[(string)boletoLog["boleto"]["id"]] = boletoLog;
                        }

                    } while (logsCursor != null);


                    logsCursor = "";

                    logsParam["types"] = "registered";

                    do
                    {
                        try
                        {
                            respJson = Boleto.Log.Get(logsCursor, logsParam);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Close();
                            return;
                        }

                        boletoLogs = (JArray)respJson["logs"];
                        if ((string)respJson["cursor"] != "") logsCursor = (string)respJson["cursor"];

                        foreach (JObject boletoLog in boletoLogs)
                        {
                            logsRegisteredByCharge[(string)boletoLog["boleto"]["id"]] = boletoLog;
                        }
                    } while (logsCursor != null);

                    foreach (JObject boleto in boletos)
                    {
                        if ((string)boleto["status"] == "paid")
                        {
                            SetBoletoInfo(boleto, (JObject)logsPaidByCharge[(string)boleto["id"]], (JObject)logsRegisteredByCharge[(string)boleto["id"]], logRow);
                        }

                        logRow++;
                    }

                    logsPaidByCharge = new Dictionary<string, object>();
                    logsRegisteredByCharge = new Dictionary<string, object>();
                }

            } while (cursor != null);

            Close();
        }

        private void SetBoletoInfo(JObject boleto, JObject paidLog, JObject createdLog, int row)
        {
            var worksheet = Globals.GetBoleto;

            double amount = double.Parse((string)boleto["amount"]) / 100;
            worksheet.Range["J" + row].Value = new StarkDateTime((string)paidLog["created"]).ToString();

            double nominalAmount = amount;
            if (createdLog["boleto"] != null) nominalAmount = double.Parse((string)createdLog["boleto"]["amount"]) / 100;

            double deltaAmount = amount - nominalAmount;

            worksheet.Range["F" + row].Value = nominalAmount;

            if (deltaAmount < 0)
            {
                Range discountCell = worksheet.Range["G" + row];
                discountCell.Value = deltaAmount;
                discountCell.Font.Color = XlRgbColor.rgbGreen;
            }
            if (deltaAmount > 0)
            {
                var interest = amount - (nominalAmount * (1 + (double)boleto["fine"] / 100));
                var fine = amount - (interest + nominalAmount);

                Range fineCell = worksheet.Range["H" + row];
                fineCell.Value = fine;
                fineCell.Font.Color = XlRgbColor.rgbRed;

                Range interestCell = worksheet.Range["I" + row];
                interestCell.Value = interest;
                interestCell.Font.Color = XlRgbColor.rgbRed;
            }
        }

        private string GetStatus(string status)
        {
            switch (status)
            {
                case "Todos":
                    return "all";
                case "Pagos":
                    return "paid";
                case "Pendentes de Registro":
                    return "created";
                case "Registrados":
                    return "registered";
                case "Vencidos":
                    return "overdue";
                case "Cancelados":
                    return "canceled";
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
                case "created":
                    return "pendente de registro";
                case "registered":
                    return "registrado";
                case "overdue":
                    return "vencido";
                case "canceled":
                    return "cancelado";
                case "failed":
                    return "falha";
                case "unknown":
                    return "desconhecido";
                default:
                    return "status inválido";
            }
        }

        private void OptionButtonEventCredited_CheckedChanged(object sender, EventArgs e)
        {
            if (OptionButtonEventCredited.Checked)
            {
                OptionButtonEventCredited.Checked = true;
            }
            else
            {
                OptionButtonEventCredited.Checked = false;
            }
        }
    }
}
