using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ApiClient.Models;
using ApiClient.OAuth2;
using ApiClient;
using Common.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Threading;

namespace API_Digi_Key_Parser_new
{
    public partial class Form1 : Form
    {
        ApiClientSettings settings;
        ApiClientService client;
        int count = 0;
        private static readonly ILog _log = LogManager.GetLogger(typeof(Program));

        public Form1()
        {
            InitializeComponent();
            ParserInit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            List<string> InputData = new List<string>();
            ConnectToExcel ConnectToExcel = new ConnectToExcel(@"D:\TestExcel.xlsx");
            ListOfPartNumbers ListOfPartNumbers = new ListOfPartNumbers(@"D:\TestExcel.xlsx");
            InputData = ListOfPartNumbers.GetListOfPartNumbers();

            for (int i = 0; i < InputData.Count; i++)
            {
                Task task = FindPartNumbers(InputData[i]);
                await task;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {

        }
        async void ParserInit()
        {
            try
            {
                settings = ApiClientSettings.CreateFromConfigFile();
                client = new ApiClientService(settings);
                if (settings.ExpirationDateTime < DateTime.Now)
                {
                    // Let's refresh the token
                    var oAuth2Service = new OAuth2Service(settings);
                    var oAuth2AccessToken = await oAuth2Service.RefreshTokenAsync();
                    if (oAuth2AccessToken.IsError)
                    {
                        // Current Refresh token is invalid or expired 
                        _log.DebugFormat("Current Refresh token is invalid or expired ");
                        textBox1.AppendText("Current Refresh token is invalid or expired " + Environment.NewLine);
                        return;
                    }

                    settings.UpdateAndSave(oAuth2AccessToken);

                    _log.DebugFormat("After call to refresh");
                    _log.DebugFormat(settings.ToString());

                    textBox1.AppendText("After call to refresh" + Environment.NewLine);
                    textBox1.AppendText(settings.ToString());
                }
            }
            catch (Exception)
            {
                textBox1.AppendText("Exception...");
                throw;
            }
        }
        async Task FindPartNumbers(string PartNumber)
        {
            try
            { 
                var response = await client.KeywordSearch(PartNumber);

                // In order to pretty print the json object we need to do the following
                var jsonFormatted = JToken.Parse(response).ToString(Formatting.Indented);

                int start = jsonFormatted.IndexOf("\"Value\": ");
                int end = jsonFormatted.IndexOf('}');

                textBox1.AppendText($"Reponse is {jsonFormatted.Substring(start, end - start)}" + Environment.NewLine);
            }
            catch (Exception)
            {
                textBox1.AppendText("Exception...");
                throw;
            }
        }
    }
}
