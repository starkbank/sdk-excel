using System;
using System.IO;
using System.Text;
using EllipticCurve;
using System.Numerics;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace StarkBankExcel.Resources
{
    public class B64ToFile
    {
        public static bool b64ToFile(string attachmentString)
        {
            // string attachment = attachmentString.Substring(attachmentString.IndexOf("base64,") + "base64,".Length);
            // string[] parts = attachmentString.Split(new[] { ";base64," }, StringSplitOptions.None);

            if (true)
            {

                // string contentType = parts[0].Split(':')[1];
                // string extension = contentType.Split('/')[1];

                // string contentType = attachmentString

                string extension = "png";

                string fileName = "qrcode-starkbank";
                string selectedPath = "C:\\Users\\Stark - Admin\\Documents\\qrcode";

                byte[] attachmentb64 = Convert.FromBase64String(attachmentString);

                // string fileName = worksheet.Range["A" + i].Value.ToString().Substring(0, 10).Replace("/", "") + "-" + worksheet.Range["B" + i].Value + "-" + worksheet.Range["D" + i].Value;
                // fileName = Regex.Replace(fileName, "[*|@|*|&]", string.Empty);

                File.WriteAllBytes(selectedPath + "\\" + fileName + "." + extension, attachmentb64);

                return true;
            }
        }

    }
}
