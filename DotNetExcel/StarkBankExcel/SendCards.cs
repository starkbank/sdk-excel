using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using StarkBankExcel.Forms;

namespace StarkBankExcel
{
    public partial class SendCards
    {
        private void Planilha18_Startup(object sender, System.EventArgs e)
        {
        }

        private void Planilha18_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button3.Click += new System.EventHandler(this.button3_Click);
            this.button4.Click += new System.EventHandler(this.button4_Click);
            this.button5.Click += new System.EventHandler(this.button5_Click);
            this.Startup += new System.EventHandler(this.Planilha18_Startup);
            this.Shutdown += new System.EventHandler(this.Planilha18_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var worksheet = Globals.SendCards;

            List<string> emptyList = new List<string>();
            Dictionary<string, object> body = new Dictionary<string, object>
            {
                { "tags", emptyList }
            };

            JObject fetchedJson;
            JObject kitIdObjects;

            fetchedJson = Request.Fetch(
                Request.Post,
                Globals.Credentials.Range["B3"].Value,
                "corporate-shop-cart",
                body
            ).ToJson();

            kitIdObjects = Request.Fetch(
                Request.Get,
                Globals.Credentials.Range["B3"].Value,
                "corporate-shop-kit?status=active"
            ).ToJson();

            bool anySent = false;

            string returnMessage = "";
            string warningMessage = "";
            string errorMessage = "";

            var initRow = TableFormat.HeaderRow + 1;
            int lastRow = worksheet.Cells[worksheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;

            int batchSize = 100;
            int errorNum = 10;
            int batchCount = (int)Math.Ceiling((double)(lastRow - 10) / batchSize);

            if (batchCount <= 1)
            {
                batchCount = 2;
            }

            string id = (string)fetchedJson["cart"]["id"];

            Parallel.For(0, batchCount, batchIndex =>
            {
                int start = 10 + batchIndex * batchSize;
                int end = Math.Min((start + batchSize) - 1, lastRow);

                List<Dictionary<string, object>> orders = new List<Dictionary<string, object>>();
                List<int> orderNumbers = new List<int>();

                for (int row = start; row <= end; row++)
                {
                    string cartId = id;
                    string kitId = (string)kitIdObjects["kits"][0]["id"];
                    string displayName2 = worksheet.Range["B" + row].Value?.ToString();
                    string displayName1 = worksheet.Range["C" + row].Value?.ToString();
                    string holderName = worksheet.Range["C" + row].Value?.ToString();
                    string shippingPhone = worksheet.Range["D" + row].Value?.ToString();
                    string shippingStreetLine1 = worksheet.Range["E" + row].Value?.ToString();
                    string shippingStreetLine2 = worksheet.Range["F" + row].Value?.ToString();
                    string shippingDistrict = worksheet.Range["G" + row].Value?.ToString();
                    string shippingCity = worksheet.Range["H" + row].Value?.ToString();
                    string shippingStateCode = worksheet.Range["I" + row].Value?.ToString();
                    string shippingZipCode = worksheet.Range["J" + row].Value?.ToString();
                    string shippingCountryCode = "BRA";

                    if (displayName1 == null)
                    {
                        MessageBox.Show("Por favor, não deixe o campo 'Portador' em branco", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (shippingPhone == null)
                    {
                        MessageBox.Show("Por favor, não deixe o campo 'Telefone' em branco", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (shippingPhone.Trim().Substring(0, 1) != "(")
                    {
                        MessageBox.Show("Telefone deve ser enviado nesse formato: (xx) xxxxx-xxxx", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (shippingDistrict == null)
                    {
                        MessageBox.Show("Por favor, não deixe o campo 'Bairro' em branco", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (shippingStateCode == null)
                    {
                        MessageBox.Show("Por favor, não deixe o campo 'UF' em branco", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (shippingZipCode == null)
                    {
                        MessageBox.Show("Por favor, não deixe o campo 'CEP' em branco", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    Dictionary<string, object> item = new Dictionary<string, object> {
                        { "kitId", kitId },
                        { "cartId", cartId },
                        { "displayName1", displayName1 },
                        { "displayName2", displayName2 },
                        { "holderName", holderName },
                        { "shippingStreetLine1", shippingStreetLine1 },
                        { "shippingStreetLine2", shippingStreetLine2 },
                        { "shippingDistrict", shippingDistrict },
                        { "shippingCity", shippingCity },
                        { "shippingStateCode", shippingStateCode },
                        { "shippingZipCode", shippingZipCode },
                        { "shippingCountryCode", shippingCountryCode },
                        { "shippingPhone", "+55 " + shippingPhone },
                    };

                    orderNumbers.Add(row);
                    orders.Add(item);
                }

                if (orderNumbers.Count > 0)
                {
                    try
                    {
                        DateTime currentTime = DateTime.UtcNow;

                        Dictionary<string, object> payload = new Dictionary<string, object>
                        {
                            { "items", orders }
                        };

                        JObject res = Request.Fetch(
                            Request.Post,
                            Globals.Credentials.Range["B3"].Value,
                            "corporate-shop-item",
                            payload
                        ).ToJson();
                        anySent = true;

                        Globals.Credentials.Range["C6"].Value = id;
                    }
                    catch (Exception ex)
                    {
                        errorMessage = Utils.ParsingErrors(ex.Message, errorNum);
                    }
                    errorNum += 100;
                }
            });

            if (anySent)
            {
                Redirect redirect = new Redirect();
                redirect.ShowDialog();
                return;
            }

            if (!string.IsNullOrEmpty(errorMessage))
            {
                MessageBox.Show(errorMessage);
                return;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Globals.Main.Activate();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            LoginForm loginForm = new LoginForm();
            loginForm.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Utils.LogOut();
        }
    }
}
