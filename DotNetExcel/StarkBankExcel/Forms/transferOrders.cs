﻿using System;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace StarkBankExcel
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
            var worksheet = Globals.Transfers;

            label2.Visible = true;

            string teamId = comboBox1.SelectedItem.ToString();
            string pattern = @"(?<=id\s=\s)\d+";

            teamId = Regex.Match(teamId, pattern).Value;

            bool anySent = false;

            string returnMessage = "";
            string warningMessage = "";
            string errorMessage = "";

            var initRow = TableFormat.HeaderRow + 1;
            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            int batchSize = 100;
            int batchCount = (int)Math.Ceiling((double)(lastRow - 10) / batchSize);

            if (batchCount <= 1)
            {
                batchCount = 2;
            }

            Parallel.For(0, batchCount, batchIndex =>
            {
                int start = 10 + batchIndex * batchSize;
                int end = Math.Min((start + batchSize) - 1, lastRow);

                List<Dictionary<string, object>> orders = new List<Dictionary<string, object>>();
                List<int> orderNumbers = new List<int>();
                List<string> externalIds = new List<string>();

                for (int row = start; row <= end; row++)
                {
                    string name = worksheet.Range["A" + row].Value?.ToString();
                    string taxID = worksheet.Range["B" + row].Value?.ToString();
                    string amountString = worksheet.Range["C" + row].Value?.ToString();
                    if (amountString == null)
                    {
                        MessageBox.Show("Por favor, não deixe linhas em branco entre as ordens de transferência", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Close();
                        return;
                    }

                    int amount = Convert.ToInt32(double.Parse(amountString) * 100);
                    string ispb = worksheet.Range["D" + row].Value?.ToString();
                    string branchCode = worksheet.Range["E" + row].Value?.ToString();
                    string accountNumber = worksheet.Range["F" + row].Value?.ToString();
                    string due = worksheet.Range["G" + row].Value?.ToString();
                    string accountType = worksheet.Range["H" + row].Value?.ToString();
                    string tags = worksheet.Range["I" + row].Value?.ToString();
                    string description = worksheet.Range["J" + row].Value?.ToString();
                    string externalID = worksheet.Range["K" + row].Value?.ToString();

                    string calculatedExternalID = Utils.calculateExtrenalId(amount, name, taxID, ispb, branchCode, accountNumber);

                    if (calculatedExternalID == externalID)
                    {
                        warningMessage = "Aviso: Pedidos já enviados hoje não foram reenviados! \n \n";
                        MessageBox.Show(warningMessage);
                    }
                    if (calculatedExternalID != externalID)
                    {
                        Dictionary<string, object> payment = new Dictionary<string, object> {
                            { "amount", amount },
                            { "taxId", taxID },
                            { "name", name },
                            { "bankCode", ispb },
                            { "branchCode", branchCode },
                            { "accountNumber", accountNumber },
                            { "accountType", accountType }
                        };
                        orderNumbers.Add(row);
                        externalIds.Add(calculatedExternalID);

                        Dictionary<string, object> paymentRequestData = new Dictionary<string, object>();

                        if (description != null)
                        {
                            paymentRequestData = new Dictionary<string, object> {
                                { "centerId", teamId },
                                { "type", "transfer" },
                                { "payment", payment },
                                { "description", description }
                            };
                        }

                        if (description == null)
                        {
                            paymentRequestData = new Dictionary<string, object> {
                                { "centerId", teamId },
                                { "type", "transfer" },
                                { "payment", payment }
                            };
                        }

                        if (due != null) { paymentRequestData.Add("due", new StarkDate(due).ToString());}

                        if (tags != null) { paymentRequestData.Add("tags", tags.Split(','));}
                        orders.Add(paymentRequestData);

                    }
                }

                if (orderNumbers.Count > 0)
                {
                    try
                    {
                        DateTime currentTime = DateTime.UtcNow;

                        JObject res = PaymentRequest.Create(orders);
                        anySent = true;

                        string createOrder = (string)res["message"];
                        returnMessage = returnMessage + Utils.rowsMessage(start, end) + createOrder + "\n";
                        for (int j = 0; j < externalIds.Count; j++)
                        {
                            worksheet.Range["K" + orderNumbers[j]].Value = externalIds[j];
                        }
                    }
                    catch (Exception ex)
                    {
                        errorMessage = Utils.ParsingErrors(ex.Message, 10 + (100 * batchIndex));
                    }
                }
            });

            label2.Visible = false;

            if (anySent)
            {
                MessageBox.Show(warningMessage + returnMessage + errorMessage);
                return;
            }
            if(!string.IsNullOrEmpty(errorMessage))
            {
                MessageBox.Show(errorMessage);
                return;
            }
        }
    }
}
