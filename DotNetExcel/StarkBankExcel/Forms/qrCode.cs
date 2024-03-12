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
           string id = Globals.Credentials.Range["B15"].Value;
           string challengePk = Globals.Credentials.Range["B16"].Value;

            Debug.WriteLine("---b64---");
            Debug.WriteLine(bString);
            Debug.WriteLine("---end b64---");

            

            string attachment = bString.Substring(bString.IndexOf("base64,") + "base64,".Length);

           byte[] attachmentb64 = Convert.FromBase64String(attachment);

           byte[] resizedImageBytes = Resources.B64ToFile.ResizeImage(attachmentb64, 200, 200);

           Image originalImage = ByteArrayToImage(resizedImageBytes);

           pictureBox1.Image = originalImage;

           string path = "challenge/" + id;
           
            for (int i = 0; i < 100; i++)
            {
               Response response = Request.Fetch(
                  Request.Get,
                  "sandbox",
                  path,
                  null,
                  null,
                  challengePk
               );

                // Debug
                Debug.WriteLine("---challenge---");
                Debug.WriteLine(response.ToJson());
                Debug.WriteLine("---end challenge---");
            }


        }

        public string polling()
        {


            return "";
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
