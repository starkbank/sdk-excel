using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using System.Threading.Tasks;
using StarkBankExcel.Resources;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace StarkBankExcel.Forms
{
    public partial class ViewPaymentApprove : Form
    {
        public ViewPaymentApprove()
        {
            InitializeComponent();
        }

        private void ViewPaymentApprovement_Load(object sender, EventArgs e)
        {
            JObject costCenter;

            label2.Visible = false;

            Dictionary<string, object> query = new Dictionary<string, object>() { { "fields", "id, name, badgeCount" } };

            try
            {
                costCenter = CostCenter.Get(null, query);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Erro na Requisiçâo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                return;
            }

            var teams = costCenter["centers"];

            foreach (JObject t in teams)
            {
                comboBox2.Items.Add(t["name"].ToString() + " ( id = " + t["id"] + " )");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.GetPaymentApprove;

            string teamId = comboBox2.SelectedItem.ToString();
            string pattern = @"(?<=id\s=\s)\d+";

            teamId = Regex.Match(teamId, pattern).Value;

            string afterString = afterInput.Text;
            string after = new StarkDate(afterString).ToString();
            string beforeString = beforeInput.Text;
            string before = new StarkDate(beforeString).ToString();

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = worksheet.Range["A" + TableFormat.HeaderRow + ":V" + lastRow];
            range.ClearContents();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Data da Solicitação";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "Tipo de Pagamento";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Descrição";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Valor";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Solicitado Por";
            worksheet.Range["F" + TableFormat.HeaderRow].Value = "Status";
            worksheet.Range["G" + TableFormat.HeaderRow].Value = "ID do pagamento";
            worksheet.Range["H" + TableFormat.HeaderRow].Value = "Tags";

            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Nome";
            worksheet.Range["J" + TableFormat.HeaderRow].Value = "CPF/CNPJ";
            worksheet.Range["K" + TableFormat.HeaderRow].Value = "Codigo do Banco/ISPB";
            worksheet.Range["L" + TableFormat.HeaderRow].Value = "Agencia";
            worksheet.Range["M" + TableFormat.HeaderRow].Value = "Conta";

            Dictionary<string, object> optionalParam = new Dictionary<string, object>();

            string events = "";
            string types = "";

            if (afterInput.Enabled == true) optionalParam["after"] = after;
            if (beforeInput.Enabled == true) optionalParam["before"] = before;

            if (OptionButtonPendente.Checked) events += "pending";
            if (OptionButtonAgendado.Checked) events += "scheduled";
            if (OptionButtonNegado.Checked) events += "denied";

            if (radioButton3.Checked) types += "transfer";
            if (radioButton1.Checked) types += "boleto-payment";
            if (radioButton4.Checked) types += "tax-payment";
            if (radioButton2.Checked) types += "utility-payment";
            if (radioButton5.Checked) types += "brcode-payment";

            optionalParam.Add("status", events);
            optionalParam.Add("type", types);
            optionalParam.Add("centerId", teamId);

            int row = TableFormat.HeaderRow + 1;

            string cursor = "";

            Dictionary<string, object> returnedData = new Dictionary<string, object>();

            if (afterInput.Enabled == true && beforeInput.Enabled == true)
            {

                do
                {
                    JObject respJson;
                    try
                    {
                        respJson = PaymentRequest.Get(cursor, optionalParam);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        respJson = new JObject();
                        return;
                    }

                    if ((string)respJson["cursor"] != "") cursor = (string)respJson["cursor"];

                    JArray payments = (JArray)respJson["requests"];

                    foreach (JObject payment in payments)
                    {
                        worksheet.Range["A" + row].Value = new StarkDateTime((string)payment["created"]).Value;
                        worksheet.Range["B" + row].Value = payment["type"];
                        worksheet.Range["C" + row].Value = payment["description"];
                        worksheet.Range["D" + row].Value = double.Parse((string)payment["amount"]) / 100;
                        worksheet.Range["E" + row].Value = payment["actions"][1]["name"];
                        worksheet.Range["F" + row].Value = payment["status"];
                        worksheet.Range["G" + row].Value = payment["id"];
                        worksheet.Range["H" + row].Value = Utils.ListToString(payment["tags"].ToObject<List<string>>(), ",");

                        if (payment["type"].ToString() == "transfer")
                        {
                            worksheet.Range["I" + row].Value = payment["payment"]["name"];
                            worksheet.Range["J" + row].Value = payment["payment"]["taxId"];
                            worksheet.Range["K" + row].Value = payment["payment"]["bankCode"];
                            worksheet.Range["L" + row].Value = payment["payment"]["branchCode"];
                            worksheet.Range["M" + row].Value = payment["payment"]["accountNumber"];
                        }
                        if (payment["type"].ToString() == "boleto-payment")
                        {
                            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Linha Digitável / Código de Barras";
                            worksheet.Range["K" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["L" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["M" + TableFormat.HeaderRow].Value = "";

                            worksheet.Range["J" + row].Value = payment["payment"]["taxId"];

                            if (payment["payment"]["line"] != null)
                            {
                                worksheet.Range["I" + row].Value = payment["payment"]["line"];
                            }
                            if (payment["payment"]["barCode"] != null)
                            {
                                worksheet.Range["I" + row].Value = payment["payment"]["barCode"];
                            }
                        }
                        if (payment["type"].ToString() == "utility-payment")
                        {
                            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Linha Digitável / Código de Barras";
                            worksheet.Range["K" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["L" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["M" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["J" + TableFormat.HeaderRow].Value = "";

                            if (payment["payment"]["line"] != null)
                            {
                                worksheet.Range["I" + row].Value = payment["payment"]["line"];
                            }
                            if (payment["payment"]["barCode"] != null)
                            {
                                worksheet.Range["I" + row].Value = payment["payment"]["barCode"];
                            }
                        }
                        if (payment["type"].ToString() == "tax-payment")
                        {
                            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Linha Digitável / Código de Barras";
                            worksheet.Range["K" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["L" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["M" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["J" + TableFormat.HeaderRow].Value = "";

                            if (payment["payment"]["line"] != null)
                            {
                                worksheet.Range["I" + row].Value = payment["payment"]["line"];
                            }
                            if (payment["payment"]["barCode"] != null)
                            {
                                worksheet.Range["I" + row].Value = payment["payment"]["barCode"];
                            }
                        }
                        if (payment["type"].ToString() == "brcode-payment")
                        {
                            worksheet.Range["I" + TableFormat.HeaderRow].Value = "Linha Digitável / Código de Barras";
                            worksheet.Range["K" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["L" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["M" + TableFormat.HeaderRow].Value = "";
                            worksheet.Range["J" + row].Value = payment["payment"]["taxId"];

                            if (payment["payment"]["line"] != null)
                            {
                                worksheet.Range["I" + row].Value = payment["payment"]["line"];
                            }
                            if (payment["payment"]["barCode"] != null)
                            {
                                worksheet.Range["I" + row].Value = payment["payment"]["barCode"];
                            }
                        }

                        row++;
                    }
                } while (cursor != null);

                Close();
                return;
            }

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
