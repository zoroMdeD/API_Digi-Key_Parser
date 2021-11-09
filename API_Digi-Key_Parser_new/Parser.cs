using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ApiClient.Models;
using ApiClient.OAuth2;
using ApiClient;
using Common.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Threading;

namespace API_Digi_Key_Parser_new
{
    public class Parser
    {
        private string partNumber = string.Empty;
        private ApiClientSettings settings;
        private ApiClientService client;

        public string PartNumber
        {
            get
            {
                return partNumber;
            }
            private set
            {
                partNumber = value;
            }
        }

        public Parser()
        {

        }

        public async Task ParserInit()
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
                        return;
                    }

                    settings.UpdateAndSave(oAuth2AccessToken);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public async Task<string> FindDescriprions(string PartNumber)
        {
            try
            {
                var response = await client.KeywordSearch(PartNumber);

                // In order to pretty print the json object we need to do the following
                var jsonFormatted = JToken.Parse(response).ToString(Formatting.Indented);

                int start = jsonFormatted.IndexOf("\"Value\": ");
                int end = jsonFormatted.IndexOf('}');

                return $"Reponse is {jsonFormatted.Substring(start, end - start)}";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

    }
}
