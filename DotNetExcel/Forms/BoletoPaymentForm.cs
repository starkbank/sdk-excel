using System;
using System.Text;
using System.Linq;
using System.Data;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace StarkBankExcel.Forms
{
    public partial class BoletoPaymentForm : Form
    {
        public BoletoPaymentForm()
        {
            InitializeComponent();
        }

        private void BoletoPayment_Load(object sender, EventArgs e)
        {

        }
        private void Login_Click(object sender, EventArgs e)
        {

            string email = Email.Text.ToLower();
            string password = Password.Text.ToString();

            PrivateKey keys = keyGen.generateKeyFromPassword(password, email);

            string privateKeyPem = keys.toPem();

            var worksheet = Globals.BoletoPayment;

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

            var initRow = TableFormat.HeaderRow + 1;
            lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            List<int> boletoPaymentNumbers = new List<int>();

            for (int row = initRow; row <= lastRow; row++)
            {
                iteration++;

                string line = worksheet.Range["A" + row].Value?.ToString();
                string taxId = worksheet.Range["B" + row].Value?.ToString();
                string scheduled = worksheet.Range["C" + row].Value?.ToString();
                string description = worksheet.Range["D" + row].Value?.ToString();
                string tags = worksheet.Range["E" + row].Value?.ToString();

                Dictionary<string, object> boleto = new Dictionary<string, object> {
                    { "taxId", taxId },
                    { "description", description },
                };

                if (scheduled != null)
                {
                    boleto.Add("scheduled", new StarkDate(scheduled).ToString());
                }

                if (tags != null)
                {
                    boleto.Add("tags", tags.Split(','));
                }

                if (line.Split('.').Length > 1)
                {
                    boleto.Add("line", line);
                }

                if (line.Split('.').Length < 1)
                {
                    boleto.Add("barCode", line);
                }

                boletos.Add(boleto);

                boletoPaymentNumbers.Add(iteration);

                if (iteration % 100 == 0 || row >= lastRow)
                {
                    if (boletoPaymentNumbers.Count == 0) goto nextIteration;

                    try
                    {
                        JObject res = BoletoPaymentClass.Create(new Dictionary<string, object>() { { "payments", boletos } }, privateKeyPem);

                        string createoBoletoPayment = (string)res["message"];
                        returnMessage = returnMessage + Utils.rowsMessage(initRow, row) + createoBoletoPayment + "\n";
                    }
                    catch (Exception ex)
                    {
                        errorMessage = ex.Message;
                    }
                nextIteration:
                    initRow = row + 1;
                    boletoPaymentNumbers = new List<int>();
                }
            }
            MessageBox.Show(warningMessage + returnMessage + errorMessage);
        }

    }
}
