using EllipticCurve;
using StarkBankExcel.Resources;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace StarkBankExcel.Forms
{
    public partial class qrCode : Form
    {
        public qrCode()
        {
           InitializeComponent();

           string bString = Globals.Credentials.Range["B14"].Value;
           
           string email = Globals.Credentials.Range["B2"].Value;

           string attachment = bString.Substring(bString.IndexOf("base64,") + "base64,".Length);

           byte[] attachmentb64 = Convert.FromBase64String(attachment);

           byte[] resizedImageBytes = Resources.B64ToFile.ResizeImage(attachmentb64, 180, 180);

           Image originalImage = ByteArrayToImage(resizedImageBytes);

           pictureBox1.Image = originalImage;
           emailLabel.Text = email.ToString();
        }

        static Image ByteArrayToImage(byte[] byteArrayIn)
        {
            using (MemoryStream ms = new MemoryStream(byteArrayIn))
            {
                Image image = Image.FromStream(ms);
                return image;
            }
        }
    }
}
