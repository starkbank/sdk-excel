using Microsoft.Office.Tools.Excel.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StarkBankExcel.Forms
{
    public partial class qrCode3 : Form
    {
        public qrCode3()
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

        static Dictionary<string, int> locationReturner(Dictionary<string, string> inputLocation, int mxValue, int myValue)
        {

            int delta = 1000;

            int mx = mxValue;
            int my = myValue;

            // int mx = 1674;
            // int my = 2398;

            int x = Int32.Parse(inputLocation["Width"]) * delta;
            int y = Int32.Parse(inputLocation["Height"]) * delta;

            x = (x / mx);
            y = (y / my);

            Dictionary<string, int> locationDict = new Dictionary<string, int>()
            {
                { "x", x },
                { "y", y },
            };

            return locationDict;
        }

        static Dictionary<string, string> stringToDict(string input)
        {
            if (input == null) { return null; };

            input = input.Replace("{", "");
            input = input.Replace("}", "");

            string[] inputs = input.Split(',');

            Dictionary<string, string> myDict = new Dictionary<string, string>();

            foreach (string c in inputs)
            {
                string[] elements = c.Split('=');
                myDict[elements[0].Trim()] = elements[1].Trim();
            }

            return myDict;
        }
    }
}
