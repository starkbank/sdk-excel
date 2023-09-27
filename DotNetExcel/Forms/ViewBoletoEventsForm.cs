using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace StarkBankExcel.Forms
{
    public partial class ViewBoletoEventsForm : Form
    {
        public ViewBoletoEventsForm()
        {
            InitializeComponent();
        }

        private void ViewChargeEventsForm_Load(object sender, EventArgs e)
        {
            periodInput.Text = "Intervalo";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.GetBoletoEvents;

            Dictionary<string, object> logsPaidByCharge = new Dictionary<string, object>();
            Dictionary<string, object> logsRegisteredByCharge = new Dictionary<string, object>();

            string afterString = afterInput.Text;
            string after = new StarkDate(afterString).ToString();
            string beforeString = beforeInput.Text;
            string before = new StarkDate(beforeString).ToString();

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":V" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data do Evento";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "Evento";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Nome";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "CPF/CNPJ";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Valor de Emissão";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "Desconto";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Multa";
            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Juros";
            worksheet.Range["J" + TableFormat.HeaderRow].Value = "Data de Emissão";
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

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            if (afterInput.Enabled == true) optionalParam["after"] = after;
            if (beforeInput.Enabled == true) optionalParam["before"] = before;

            string events = "";

            if (OptionButtonEventCanceled.Checked) events += "canceled,";

            if (OptionButtonEventOverdue.Checked) events += "overdue,";

            if (OptionButtonEventCredited.Checked) events += "credited,";

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";

            if (events != "credited,")
            {
                optionalParam.Add("types", events);

                do
                {
                    JObject respJson;

                    try
                    {
                        respJson = Boleto.Log.Get(cursor, optionalParam);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Close();
                        return;
                    }

                    cursor = null;
                    if ((string)respJson["cursor"] != null) cursor = (string)respJson["cursor"];

                    JArray logs = (JArray)respJson["logs"];

                    foreach (JObject log in logs)
                    {
                        string logEvent = (string)log["type"];
                        string eventDate = new StarkDateTime((string)log["created"]).ToString();

                        JObject boleto = (JObject)log["boleto"];
                        string issueDate = new StarkDateTime((string)boleto["created"]).ToString();
                        string boletoStatus = (string)boleto["status"];
                        string dueDate = new StarkDateTime((string)boleto["due"]).ToString();
                        string id = (string)boleto["id"];

                        worksheet.Range["A" + row].Value = eventDate;
                        worksheet.Range["B" + row].Value = GetTypeInPt(logEvent);
                        worksheet.Range["C" + row].Value = boleto["name"];
                        worksheet.Range["D" + row].Value = boleto["taxId"];
                        worksheet.Range["E" + row].Value = double.Parse((string)boleto["amount"]) / 100;

                        worksheet.Range["J" + row].Value = issueDate;
                        worksheet.Range["K" + row].Value = dueDate;
                        worksheet.Range["L" + row].Value = boleto["line"];
                        worksheet.Range["M" + row].Value = id;
                        worksheet.Range["N" + row].Value = double.Parse((string)boleto["fee"]) / 100;
                        worksheet.Range["O" + row].Value = Utils.ListToString(boleto["tags"].ToObject<List<string>>(), ",");

                        Range rng = worksheet.Range["P" + row];
                        rng.Value = "PDF";
                        Hyperlink link = rng.Hyperlinks.Add(rng, Utils.BaseUrl(Globals.Credentials.Range["B3"].Value) + "/v2/boleto/" + id + "/pdf");
                        row++;
                    }

                } while (cursor != null);
                Close();
                return;
            }

            optionalParam.Add("types", events);

            do
            {
                int logRow = row;

                JObject respJson;

                try
                {
                    respJson = Boleto.Log.Get(cursor, optionalParam);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }

                cursor = null;
                if (respJson["cursor"] != null) cursor = (string)respJson["cursor"];

                JArray logs = (JArray)respJson["logs"];
                if (logs.Count == 0) break;

                foreach(JObject log in logs)
                {
                    string logEvent = (string)log["type"];
                    string eventDate = new StarkDateTime((string)log["created"]).ToString();

                    JObject boleto = (JObject)log["boleto"];
                    string issueDate = new StarkDateTime((string)boleto["created"]).ToString();
                    string boletoStatus = (string)boleto["status"];
                    string dueDate = new StarkDateTime((string)boleto["due"]).ToString();
                    string id = (string)boleto["id"];

                    worksheet.Range["A" + row].Value = eventDate;
                    worksheet.Range["B" + row].Value = GetTypeInPt(logEvent);
                    worksheet.Range["C" + row].Value = boleto["name"];
                    worksheet.Range["D" + row].Value = boleto["taxId"];
                    worksheet.Range["E" + row].Value = double.Parse((string)boleto["amount"]) / 100;

                    worksheet.Range["J" + row].Value = issueDate;
                    worksheet.Range["K" + row].Value = dueDate;
                    worksheet.Range["L" + row].Value = boleto["line"];
                    worksheet.Range["M" + row].Value = id;
                    worksheet.Range["N" + row].Value = double.Parse((string)boleto["fee"]) / 100;
                    worksheet.Range["O" + row].Value = Utils.ListToString(boleto["tags"].ToObject<List<string>>(), ",");

                    Range rng = worksheet.Range["P" + row];
                    rng.Value = "PDF";
                    Hyperlink link = rng.Hyperlinks.Add(rng, Utils.BaseUrl(Globals.Credentials.Range["B3"].Value) + "/v2/boleto/" + id + "/pdf");

                    worksheet.Range["Q" + row].Value = boleto["streetLine1"];
                    worksheet.Range["R" + row].Value = boleto["streetLine2"];
                    worksheet.Range["S" + row].Value = boleto["district"];
                    worksheet.Range["T" + row].Value = boleto["city"];
                    worksheet.Range["U" + row].Value = boleto["stateCode"];
                    worksheet.Range["V" + row].Value = boleto["zipCode"];

                    if (boletoStatus == "paid" && !logsPaidByCharge.ContainsKey(id))
                    {
                        logsPaidByCharge.Add(id, log);
                        logsRegisteredByCharge.Add(id, new Dictionary<string, object>());
                    }

                    row++;
                }

                Dictionary<string, object> logsParam = new Dictionary<string, object>();
                string keys = "";
                string sep = "";
                string registeredCursor = null;

                foreach (string boletoId in logsPaidByCharge.Keys)
                {
                    keys = keys + sep + boletoId;
                    sep = ",";
                }

                logsParam.Add("boletoIds", keys);

                do
                {
                    logsParam.Add("types", "registered,");

                    try
                    {
                        respJson = Boleto.Log.Get(registeredCursor, logsParam);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Close();
                        return;
                    }

                    registeredCursor = null;
                    if (respJson["cursor"] != null) registeredCursor = (string)respJson["cursor"];

                    JArray registeredLogs = (JArray)respJson["logs"];
                    foreach (JObject log in registeredLogs)
                    {
                        logsRegisteredByCharge[(string)log["boleto"]["id"]] = log;
                    }
                } while (registeredCursor != null);

                foreach(JObject log in logs)
                {
                    JObject boleto = (JObject)log["boleto"];
                    if ((string) boleto["status"] == "paid")
                    {
                        SetChargeEventInfo(boleto, (JObject) logsRegisteredByCharge[(string)boleto["id"]], logRow);
                    }
                    logRow++;
                }

                logsPaidByCharge = new Dictionary<string, object>();
                logsRegisteredByCharge = new Dictionary<string, object>();

            } while(cursor != null);

            Close();
        }

        private void SetChargeEventInfo(JObject boleto, JObject registeredLog, int row)
        {
            var worksheet = Globals.GetBoletoEvents;

            double amount = double.Parse((string)boleto["amount"]) / 100;

            double nominalAmount = amount;
            if (registeredLog["boleto"] != null) nominalAmount = double.Parse((string)registeredLog["boleto"]["amount"]) / 100;

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
                double fine = (double.Parse((string)boleto["amount"]) / 100) * nominalAmount;
                double interest = amount - fine - nominalAmount;

                Range fineCell = worksheet.Range["H" + row];
                fineCell.Value = fine;
                fineCell.Font.Color = XlRgbColor.rgbRed;

                Range interestCell = worksheet.Range["I" + row];
                interestCell.Value = interest;
                interestCell.Font.Color = XlRgbColor.rgbRed;
            }

        }

        private string GetTypeInPt(string status)
        {
            switch (status)
            {
                case "paid":
                    return "pago";
                case "bank":
                    return "creditado";
                case "register":
                    return "criado (pendente de registro)";
                case "registered":
                    return "registrado";
                case "overdue":
                    return "vencido";
                case "cancel":
                    return "em cancelamento";
                case "canceled":
                    return "cancelado";
                case "failed":
                    return "falha";
                case "unknown":
                    return "desconhecido";
                default:
                    return status;
            }
        }
    }
}
