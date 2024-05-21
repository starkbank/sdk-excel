using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace StarkBankExcel.Forms
{
    public partial class VersionWarning : Form
    {
        public VersionWarning()
        {
            InitializeComponent();;

            try
            {

                string url = "https://raw.githubusercontent.com/starkbank/sdk-excel/master/CHANGELOG.md";

                HttpRequestMessage httpRequestMessage = new HttpRequestMessage
                {
                    Method = new HttpMethod("GET"),
                    RequestUri = new Uri(url)
                };

                HttpClient Client = new HttpClient();
                Client.DefaultRequestHeaders.Add("User-Agent", "Excel-DotNet");
                httpRequestMessage.Headers.TryAddWithoutValidation("Content-Type", "application/json");
                httpRequestMessage.Headers.TryAddWithoutValidation("Accept-Language", "pt-BR");
                httpRequestMessage.Headers.TryAddWithoutValidation("Accept", "*/*");

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                var result = Client.SendAsync(httpRequestMessage).Result;
                Response response = new Response(
                result.Content.ReadAsByteArrayAsync().Result,
                (int)result.StatusCode
                );

                string versionWarning = response.Content;

                string[] separate = { "[Unreleased]" };

                versionWarning = versionWarning.Split(separate, System.StringSplitOptions.None)[1].Split('-')[0].Split('[')[1].Substring(0, 5).Trim();

                version.Text = "Versão nova: " + versionWarning.ToString();
            }
            catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string url = "https://github.com/starkbank/sdk-excel/raw/master/StarkBankInstaller.exe";
            Process.Start(url);
        }
    }
}
