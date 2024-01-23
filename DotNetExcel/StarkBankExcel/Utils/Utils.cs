using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Newtonsoft.Json.Linq;
using StarkBankExcel.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace StarkBankExcel
{
    internal class Utils
    {
        public static string calculateExtrenalId(
            int amount, string name, string taxID, string bankCode, 
            string branchCode, string accountNumber
        )
        {
            return bankCode + branchCode + accountNumber + name + taxID + amount.ToString();
        }

        public static string ParsingErrors(string element, int number)
        {
            JObject json = JObject.Parse(element);

            string result = "";
            string errorMessage = "";

            foreach (JObject errorJson in json["errors"])
            {
                try
                {
                    string error = string.Join(" ", errorJson["message"].ToString().Split(':'));
                    string splited = errorJson["message"].ToString().Split(':')[0].Substring(8);
                    int resultNumber = number + Convert.ToInt32(splited);
                    result = resultNumber.ToString();
                    errorMessage += "Linha: " + result + " Erro: " + error + "\n\n";
                }
                catch (Exception ex) { errorMessage = errorJson["message"].ToString().Split(':')[1]; }
            };

            return errorMessage;
        }

        public static string rowsMessage(int startRow, int currentRow)
        {
            return "Linhas " + startRow.ToString() + " a " + currentRow.ToString() + ": ";
        }

        public static string ListToString(List<string> list, string delimiter = null)
        {
            if (list == null || list.Count == 0) return "";

            string elString = "";
            if(list.Count != 0) 
            {
                foreach(string el in list)
                {
                    elString += el + delimiter;
                }
            }
            return elString.Substring(0, elString.Length - 1);
        }

        public static string BaseUrl(string environment)
        {
            if (environment == "production")
            {
                return "https://api.starkbank.com/";
            }
            if (environment == "sandbox")
            {
                return "https://sandbox.api.starkbank.com/";
            }
            throw new Exception("Necessario configurar ambiente");
        }

        public static string MoneyStringFrom(double value) {
            return string.Format("{0:C2}", value / 100);
        }

        public static void DisplayInfo(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            ws.Range["A2"].Value = "Olá " + Globals.Credentials.Range["B5"].Value;
            ws.Range["A3"].Value = "Workspace: " + Globals.Credentials.Range["B3"].Value;
            ws.Range["A4"].Value = "ID do Workspace: " + Globals.Credentials.Range["B6"].Value;
            ws.Range["A5"].Value = "E-mail: " + Globals.Credentials.Range["B2"].Value;
            ws.Range["A6"].Value = "Ambiente: " + Globals.Credentials.Range["B3"].Value;
            ws.Range["A7"].Value = "Saldo: " + MoneyStringFrom(Balance.Get());
        }

        public static void ClearData(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            int lastRow = ws.Cells[ws.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = ws.Range["A" + TableFormat.HeaderRow +  ":Z" + lastRow];
            range.ClearContents();
        }

        public static void ClearInfo(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            ws.Range["A2"].Value = "Olá ";
            ws.Range["A3"].Value = "Workspace: ";
            ws.Range["A4"].Value = "ID do Workspace: ";
            ws.Range["A5"].Value = "E-mail: ";
            ws.Range["A6"].Value = "Ambiente: ";
            ws.Range["A7"].Value = "Saldo: ";
        }

        public static void ClearAll(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            int lastRow = ws.Cells[ws.Rows.Count, "A"].End[XlDirection.xlUp].Row;
            Range range = ws.Range["A" + (TableFormat.HeaderRow + 1) + lastRow];
            range.ClearContents();
        }

        public static void LogOut()
        {
            string message1 = "Você quer mesmo encerrar a sessâo? ";
            string message2 = "Dado que nâo foram salvos serâo apagados.";
            string confirmationMessage = message1 + message2;

            DialogResult signOutAnswer = MessageBox.Show(confirmationMessage, "Confirmação de encerramento", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (signOutAnswer == DialogResult.No)
            {
                return;
            }

            Session.Delete();

            Microsoft.Office.Interop.Excel.Workbook workbook = Globals.ThisWorkbook.Application.ActiveWorkbook;

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name != "Credentials")
                {
                    ClearAll(sheet);
                    ClearInfo(sheet);
                }
            }
        }
    }
}
