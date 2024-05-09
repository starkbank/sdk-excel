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

            string formSize = Globals.Credentials.Range["C16"].Value;

            Dictionary<string, string> sizeValue = new Dictionary<string, string>();

            sizeValue = stringToDict(formSize);

            if (sizeValue != null)
            {
                Dictionary<string, int> qrCodeLocationValue = locationReturner(sizeValue, 1620, 2398);

                Dictionary<string, int> textLocationValue = locationReturner(sizeValue, 10880, 4567);

                Dictionary<string, System.Single> textSizeValue = sizeReturner(sizeValue, 32);

                string bString = Globals.Credentials.Range["B14"].Value;

                string email = Globals.Credentials.Range["B2"].Value;

                string attachment = bString.Substring(bString.IndexOf("base64,") + "base64,".Length);

                byte[] attachmentb64 = Convert.FromBase64String(attachment);

                byte[] resizedImageBytes = Resources.B64ToFile.ResizeImage(attachmentb64, 180, 180);

                Image originalImage = ByteArrayToImage(resizedImageBytes);

                pictureBox1.Image = originalImage;
                pictureBox2.Location = new System.Drawing.Point(0, 0);
                pictureBox2.Size = new Size(Int32.Parse(sizeValue["Width"]), Int32.Parse(sizeValue["Height"]));
                pictureBox1.Location = new System.Drawing.Point(qrCodeLocationValue["x"], qrCodeLocationValue["y"]);
                emailLabel.Location = new System.Drawing.Point(textLocationValue["x"], textLocationValue["y"]);
                emailLabel.Font = new System.Drawing.Font("Mundial", textSizeValue["font"], System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

                emailLabel.Text = email.ToString();
            }
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

        static Dictionary<string, System.Single> sizeReturner(Dictionary<string, string> inputLocation, System.Single value)
        {

            value = value / 10;

            int areaValue = Int32.Parse(inputLocation["Width"]) * Int32.Parse(inputLocation["Height"]);

            System.Single pointValue = (areaValue * value) / 100000;

            Dictionary<string, System.Single> locationDict = new Dictionary<string, System.Single>()
            {
                { "font", pointValue },
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
