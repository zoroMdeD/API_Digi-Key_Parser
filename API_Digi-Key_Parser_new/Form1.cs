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

namespace API_Digi_Key_Parser_new
{
    public partial class Form1 : Form
    {
        private static readonly ILog _log = LogManager.GetLogger(typeof(Program));

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            var settings = ApiClientSettings.CreateFromConfigFile();
            _log.DebugFormat(settings.ToString());
            textBox1.AppendText(settings.ToString());
            try
            {
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

                var client = new ApiClientService(settings);
                var response = await client.KeywordSearch("stm32f407vet6");    

                // In order to pretty print the json object we need to do the following
                var jsonFormatted = JToken.Parse(response).ToString(Formatting.Indented);

                textBox1.AppendText($"Reponse is {jsonFormatted} ");
            }
            catch (Exception)
            {
                textBox1.AppendText("Exception...");
                throw;
            }

        }
    }
}
