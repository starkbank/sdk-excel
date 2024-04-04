using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StarkBankExcel.Forms
{
    public partial class Redirect : Form
    {
        public Redirect()
        {
            InitializeComponent();
        }

        private void Login_Click(object sender, EventArgs e)
        {
            string workspace = Globals.Credentials.Range["B1"].Value;
            string environment = Globals.Credentials.Range["B3"].Value;
            string cartId = Globals.Credentials.Range["C6"].Value;

            string url = "https://" + workspace + "." + environment + ".starkbank.com/cart/" + cartId;
            Process.Start(url);
        }
    }
}
