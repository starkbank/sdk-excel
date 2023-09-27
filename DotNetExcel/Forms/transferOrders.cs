using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace StarkBankMVP
{
    public partial class transferOrders : Form
    {
        public transferOrders()
        {
            InitializeComponent();
        }

        private void transferOrders_Load(object sender, EventArgs e)
        {
            JObject costCenter;

            try
            {
                 costCenter = CostCenter.Get();
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
            var worksheet = Globals.Transfers;

            string teamId = comboBox1.SelectedItem.ToString();
            string pattern = @"(?<=id\s=\s)\d+";

            teamId = Regex.Match(teamId, pattern).Value;

            List<int> orderNumbers = new List<int>();
            List<string> externalIds = new List<string>();

            bool anySent = false;

            string returnMessage = "";
            string warningMessage = "";
            string errorMessage = "";


            int iteration = 0;

            var initRow = TableFormat.HeaderRow + 1;
            var lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            List<Dictionary<string, object>> orders = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> jsonData = new List<Dictionary<string, object>>();

            for (int row = initRow; row <= lastRow; row++)
            {
                iteration++;

                string name = worksheet.Range["A" + row].Value?.ToString();
                string taxID = worksheet.Range["B" + row].Value?.ToString();
                string amountString = worksheet.Range["C" + row].Value?.ToString();
                if(amountString == null)
                {
                    MessageBox.Show("Por favor, não deixe linhas em branco entre as ordens de transferência", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }
                int amount = int.Parse(amountString);
                string ispb = worksheet.Range["D" + row].Value?.ToString();
                string branchCode = worksheet.Range["E" + row].Value?.ToString();
                string accountNumber = worksheet.Range["F" + row].Value?.ToString();
                string accountType = worksheet.Range["G" + row].Value?.ToString();

                string tags = worksheet.Range["H" + row].Value?.ToString();
                string description = worksheet.Range["I" + row].Value?.ToString();
                string externalID = worksheet.Range["J" + row].Value?.ToString();

                string calculatedExternalID = Utils.calculateExtrenalId(amount, name, taxID, ispb, branchCode, accountNumber);

                if(calculatedExternalID == externalID)
                {
                    warningMessage = "Aviso: Pedidos já enviados hoje não foram reenviados! \n \n";
                }
                else
                {
                    Dictionary<string, object> payment = new Dictionary<string, object> {
                        {"amount", amount },
                        {"taxId", taxID },
                        {"name", name},
                        {"bankCode", ispb },
                        {"branchCode", branchCode },
                        {"accountNumber", accountNumber },
                        {"accountType", accountType }
                    };

                    orderNumbers.Add(iteration);
                    externalIds.Add(calculatedExternalID);
                    
                    if (tags != null) { payment["tags"] = new List<string> { tags }; }

                    if (description != null) { payment["description"] = description; }

                    orders.Add(new Dictionary<string, object> {
                        {"centerId", teamId },
                        {"type", "transfer" },
                        {"payment", payment }
                    });
                }

                if(iteration % 100 == 0 || row >= lastRow)
                {
                    if(orderNumbers.Count == 0) goto nextIteration;


                    try
                    {
                        DateTime currentTime = DateTime.UtcNow;

                        jsonData.Add(
                            new Dictionary<string, object>
                            {
                            {"id", row },
                            {"externalId", externalID },
                            {"Time", currentTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") }
                            }
                        );

                        JObject res = PaymentRequest.Create(orders);
                        anySent = true;

                        string createOrder = (string)res["message"];
                        returnMessage = returnMessage + Utils.rowsMessage(initRow, row) + createOrder + "\n";
                        for (int j = 0; j < externalIds.Count; j++)
                        {
                            worksheet.Range["J" + (TableFormat.HeaderRow + orderNumbers[j])].Value = externalIds[j];
                        }
                    }
                    catch(Exception ex)
                    {
                        errorMessage = ex.Message;
                    }
            nextIteration:
                initRow = row + 1;
                orders = new List<Dictionary<string, object>>();
                orderNumbers = new List<int>();
                externalIds = new List<string>();
                }
            }

            string jsonString = Json.Encode(jsonData);

            File.WriteAllText(@"C:\Users\Stark - Admin\Documents\transferC#.json", jsonString);


            if (anySent)
            {
                MessageBox.Show(warningMessage + returnMessage + errorMessage);
                return;
            }
            MessageBox.Show("Todos os pedidos listados já foram enviados");
        }
    }
}
