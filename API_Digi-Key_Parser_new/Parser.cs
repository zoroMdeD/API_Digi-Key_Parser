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
using System.Linq;
using System.Text;

namespace API_Digi_Key_Parser_new
{
    public class Parser
    {
        private List<string> partNumber = new List<string>();
        private List<string> path = new List<string>();
        private List<string> getPassiveComponents = new List<string>();
        private List<string> getUniversalEquipment = new List<string>();
        Dictionary<string, string> charReplace  = new Dictionary<string, string>();
        
        private byte[] utf8Space = new byte[] { 0xC2, 0xA0 };
        private string tempSpace;

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
        public List<string> Path
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
        public List<string> GetPassiveComponents
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
        public List<string> GetUniversalEquipment
        {
            get
            {
                return getUniversalEquipment;
            }
            private set
            {
                getUniversalEquipment = value;
            }
        }

        public Parser(string pathInfoPartNumberPass, string pathInfoUniversalEquip, string pathInfoEngineers)
        {
            charReplace.Add("А", "A");
            charReplace.Add("В", "B");
            charReplace.Add("С", "C");
            charReplace.Add("Е", "E");
            charReplace.Add("Н", "H");
            charReplace.Add("К", "K");
            charReplace.Add("М", "M");
            charReplace.Add("О", "O");
            charReplace.Add("Р", "P");
            charReplace.Add("Т", "T");
            charReplace.Add("Х", "X");
            charReplace.Add("а", "a");
            charReplace.Add("с", "c");
            charReplace.Add("е", "e");
            charReplace.Add("о", "o");
            charReplace.Add("р", "p");
            charReplace.Add("х", "x");

            tempSpace = Encoding.GetEncoding("UTF-8").GetString(utf8Space);

            this.path.Add(pathInfoPartNumberPass);
            this.path.Add(pathInfoUniversalEquip);
            this.path.Add(pathInfoEngineers);
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
        private bool FindCyrillicSymbol (string keyword)
        {
            var cyrillic = Enumerable.Range(1024, 256).Select(ch => (char)ch);
            bool result = keyword.Any(cyrillic.Contains);

            return result;
        }
        private string FindSpecialSymbol(string keyword)
        {
            keyword = (keyword.Replace(" ", "")).Replace(tempSpace, "");
            //keyword = keyword.Replace(tempSpace, "");

            return keyword;
        }
        public async Task<string> FindDescriprions(string partNumber)
        {
            try
            {
                partNumber = FindSpecialSymbol(partNumber);
                if (FindCyrillicSymbol(partNumber))
                {
                    foreach (KeyValuePair<string, string> pair in charReplace)
                    {
                        partNumber = partNumber.Replace(pair.Key, pair.Value);
                    }
                }

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
                bool checkPassiveComponent = ActionWithExcel.UpdateExcelDoc(Path[0], 0, Family);    //Checking for passive components                                                                                    
                if(checkPassiveComponent)
                    getPassiveComponents.Add("Passive");
                else
                    getPassiveComponents.Add("null");
  
                getUniversalEquipment.Add(ActionWithExcel.UpdateExcelDocForReadUniversalEquipmentFile(Path[1], 0, Family));    //Checking for universal equipment


                if (!checkPassiveComponent) //Checking for passive components
                {
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
