using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace StarkBankExcel.Forms
{
    public partial class ViewHelpForm : Form
    {
        public ViewHelpForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start("https://starkbank.com/sandbox");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start("https://web.starkbank.com/signup/email");
        }
    }
}
