using System;
using System.Text;
using System.Linq;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace StarkBankExcel.Forms
{
    public partial class Redirect : Form
    {
        public Redirect()
        {
            InitializeComponent();
        }
        private void ShopClick_Click(object sender, EventArgs e)
        {
            string url = "https://" + Globals.Credentials.Range["B1"].Value + "." + Globals.Credentials.Range["B3"].Value + ".starkbank.com/cart/" + Globals.Credentials.Range["C6"].Value;
            Process.Start(url);
        }
    }
}
