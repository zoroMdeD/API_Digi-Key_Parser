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
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace API_Digi_Key_Parser_new
{
    public class Parser
    {
        private List<string> partNumber = new List<string>();
        private string path = string.Empty;
        private Dictionary<string, string> getPassiveComponents = new Dictionary<string, string>();

        private ApiClientSettings settings;
        private ApiClientService client;

        public List<string> PartNumber
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
        public string Path
        {
            get
            {
                return path;
            }
            private set
            {
                path = value;
            }
        }
        public Dictionary<string, string> GetPassiveComponents
        {
            get
            {
                return getPassiveComponents;
            }
            private set
            {
                getPassiveComponents = value;
            }
        }

        public Parser(string path)
        {
            this.path = path;
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
        public async Task<string> FindDescriprions(string partNumber)
        {
            try
            {
                this.partNumber.Add(partNumber);
                string Family;
                string Package;
                string FamilyPackage;

                var response = await client.KeywordSearch(partNumber);
                // In order to pretty print the json object we need to do the following
                var jsonFormatted = JToken.Parse(response).ToString(Formatting.Indented);

                //Find Family
                string s = "\"Value\": ";
                char[] charToTrim = { ' ', '\n', '\"', '\\', '\r' };
                int start = jsonFormatted.IndexOf(s);
                int end = jsonFormatted.IndexOf('}');

                Family = (jsonFormatted.Substring(start + s.Length, end - (start + s.Length))).Trim(charToTrim);

                //Здесь проверить на пассивку, если да то остальное не парсить, вывести Family, и прописать в столбцы Engineers, Difficult, (MotherBoard, Adapters => "PASS")
                ActionWithExcel ActionWithExcel = new ActionWithExcel();
                bool check = ActionWithExcel.UpdateExcelDoc(Path, 0, Family);

                if (!check) //Checking for passive components
                {
                    getPassiveComponents.Add(partNumber, "Not passive");
                    if (Family != "Out of Bounds")
                    {
                        //Find Package/Case
                        s = "\"Parameter\": \"Package / Case\",";
                        start = jsonFormatted.IndexOf(s);
                        end = jsonFormatted.IndexOf("\"Parameter\": \"Supplier Device Package\",");

                        Package = jsonFormatted.Substring(start + s.Length, end - (start + s.Length));

                        s = "\"Value\": ";
                        start = Package.IndexOf(s);
                        end = Package.IndexOf('}');

                        Package = (Package.Substring(start + s.Length, end - (start + s.Length))).Trim(charToTrim);

                        FamilyPackage = Family + "#" + Package;

                        return FamilyPackage;
                    }
                    else
                    {
                        return "null";
                    }
                }
                else 
                {
                    //Завести массив/список для хранения статуса на пассивность текущего партномера (запрашивать его при необходимости)
                    getPassiveComponents.Add(partNumber, "Passive");
                    return Family;
                }
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }
    }
}
