using System;
using System.IO;
using System.Text;
using EllipticCurve;
using System.Numerics;
using System.Security.Cryptography;
using System.Windows.Forms;
using System.Drawing;

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

                byte[] resizedImageBytes = ResizeImage(attachmentb64, 200, 200);

                File.WriteAllBytes(selectedPath + "\\" + fileName + "." + extension, resizedImageBytes);

                return true;
            }
        }

        public static byte[] ResizeImage(byte[] originalImageBytes, int newWidth, int newHeight)
        {
            using (MemoryStream msOriginal = new MemoryStream(originalImageBytes))
            using (Image originalImage = Image.FromStream(msOriginal))
            using (Bitmap resizedImage = new Bitmap(newWidth, newHeight))
            using (Graphics g = Graphics.FromImage(resizedImage))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(originalImage, 0, 0, newWidth, newHeight);

                using (MemoryStream msResized = new MemoryStream())
                {
                    resizedImage.Save(msResized, originalImage.RawFormat); // Preserve the original image format
                    return msResized.ToArray();
                }
            }
        }

    }
}
