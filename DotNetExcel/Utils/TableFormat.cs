using System;
using System.Linq;
using System.Text;
using StarkBankExcel;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace StarkBankExcel
{
    internal class TableFormat
    {
        public static int HeaderRow = 9;

        public static void FreezeHeader()
        {
            Microsoft.Office.Interop.Excel.Window activeWindow = Globals.ThisWorkbook.Application.ActiveWindow;

            if (activeWindow.FreezePanes)
            {
                activeWindow.FreezePanes = false;
            }

            activeWindow.SplitRow = HeaderRow;
            activeWindow.SplitColumn = 0;
            activeWindow.FreezePanes = true;

        }
    }
}
