using System;
using System.Data;
using System.Linq;
using System.Text;
using EllipticCurve;
using System.Drawing;
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
    public partial class BoletoPaymentRequest : Form
    {
        public BoletoPaymentRequest()
        {
            InitializeComponent();
        }

        private void BoletoPaymentRequest_Load(object sender, EventArgs e)
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
                comboBox1.Items.Add(t["name"].ToString() + " ( id = " + t["id"] + " )");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.BoletoPayment;

            label2.Visible = true;

            string teamId = comboBox1.SelectedItem.ToString();
            string pattern = @"(?<=id\s=\s)\d+";

            teamId = Regex.Match(teamId, pattern).Value;

            string email = Email.Text.ToLower();
            string password = Password.Text.ToString();

            PrivateKey keys = keyGen.generateKeyFromPassword(password, email);

            string privateKeyPem = keys.toPem();

            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            List<Dictionary<string, object>> boletos = new List<Dictionary<string, object>>();

            worksheet.Range["A" + TableFormat.HeaderRow].Value = "Linha Digitável ou Código de Barras";
            worksheet.Range["B" + TableFormat.HeaderRow].Value = "CPF/CNPJ do beneficiário";
            worksheet.Range["C" + TableFormat.HeaderRow].Value = "Data de Agendamento";
            worksheet.Range["D" + TableFormat.HeaderRow].Value = "Descrição";
            worksheet.Range["E" + TableFormat.HeaderRow].Value = "Tags";

            string returnMessage = "";
            string warningMessage = "";
            string errorMessage = "";

            int iteration = 0;
            int errorNum = 10;

            var initRow = TableFormat.HeaderRow + 1;
            lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            List<int> boletoPaymentNumbers = new List<int>();

            for (int row = initRow; row <= lastRow; row++)
            {
                iteration++;

                string line = worksheet.Range["A" + row].Value?.ToString();
                string taxId = worksheet.Range["B" + row].Value?.ToString();
                string due = worksheet.Range["C" + row].Value?.ToString();
                string description = worksheet.Range["D" + row].Value?.ToString();
                string tags = worksheet.Range["E" + row].Value?.ToString();

                Dictionary<string, object> payment = new Dictionary<string, object>();

                Dictionary<string, object> boleto = new Dictionary<string, object> {
                    { "taxId", taxId },
                    { "description", description },
                };

                if (line.Split('.').Length > 1)
                {
                    boleto.Add("line", line);
                }

                if (line.Split('.').Length < 1)
                {
                    boleto.Add("barCode", line);
                }

                payment.Add("centerId", teamId);
                payment.Add("type", "boleto-payment");
                payment.Add("payment", boleto);

                if (due != null)
                {
                    payment.Add("due", new StarkDate(due).ToString());
                }

                if (tags != null)
                {
                    payment.Add("tags", tags.Split(','));
                }

                boletos.Add(payment);

                boletoPaymentNumbers.Add(iteration);

                if (iteration % 100 == 0 || row >= lastRow)
                {

                    if (boletoPaymentNumbers.Count == 0) goto nextIteration;

                    try
                    {
                        JObject res = PaymentRequest.Create(boletos, privateKeyPem);

                        string createBoletoPayment = (string)res["message"];
                        returnMessage = returnMessage + Utils.rowsMessage(initRow, row) + createBoletoPayment + "\n";
                    }
                    catch (Exception ex)
                    {
                        errorMessage = Utils.ParsingErrors(ex.Message, errorNum);
                    }

                    errorNum += 100;

                nextIteration:
                    initRow = row + 1;
                    boletos = new List<Dictionary<string, object>>();
                    boletoPaymentNumbers = new List<int>();
                }
            }
            MessageBox.Show(warningMessage + returnMessage + errorMessage);


        }

        private void Email_TextChanged(object sender, EventArgs e)
        {

        }

        private void Password_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
