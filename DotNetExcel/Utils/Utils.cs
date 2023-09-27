using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StarkBankMVP
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

        public static string rowsMessage(int startRow, int currentRow)
        {
            return "Linhas " + startRow.ToString() + " a " + currentRow.ToString() + ": ";
        }

        public static string ListToString(List<string> list, string delimiter = null)
        {
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
    }
}
