﻿using System;
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
