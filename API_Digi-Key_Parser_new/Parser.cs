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
        private List<string> family = new List<string>();
        private List<string> package = new List<string>();
        private List<string> passiveComponents = new List<string>();
        private List<string> universalEquipment = new List<string>();
        private List<string> engineer = new List<string>();
        private List<int> difficulty = new List<int>();
        private List<string> motherBoard = new List<string>();
        private List<string> motherBoardTrim = new List<string>();

        Dictionary<string, string> charReplace  = new Dictionary<string, string>();

        private string subStr = string.Empty;
        private int startIndex = 0;
        private int endIndex = 0;
        private char[] charToTrim = { ' ', '\n', '\"', '\\', '\r' };
        private char[] PartNumberTrim = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
                                          'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 
                                          '_', '-' };

        private string tempSpace;

        private ApiClientSettings settings;
        private ApiClientService client;
        private ActionWithExcel ActionWithExcel;

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
        public List<string> Family
        {
            get
            {
                return family;
            }
            private set
            {
                family = value;
            }
        }
        public List<string> Package
        {
            get
            {
                return package;
            }
            private set
            {
                package = value;
            }
        }
        public List<string> PassiveComponents
        {
            get
            {
                return passiveComponents;
            }
            private set
            {
                passiveComponents = value;
            }
        }
        public List<string> UniversalEquipment
        {
            get
            {
                return universalEquipment;
            }
            private set
            {
                universalEquipment = value;
            }
        }
        public List<string> Enginner
        {
            get
            {
                return engineer;
            }
            private set
            {
                engineer = value;
            }
        }
        public List<int> Difficulty
        {
            get
            {
                return difficulty;
            }
            private set
            {
                difficulty = value;
            }
        }
        public List<string> MotherBoard
        {
            get
            {
                return motherBoard;
            }
            private set
            {
                motherBoard = value;
            }
        }
        public List<string> MotherBoardTrim
        {
            get
            {
                return motherBoardTrim;
            }
            private set
            {
                motherBoardTrim = value;
            }
        }
        public Parser()
        {
            byte[] utf8Space = new byte[] { 0xC2, 0xA0 };
            tempSpace = Encoding.GetEncoding("UTF-8").GetString(utf8Space);

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
        }
        public async Task<string> ParserInit()
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
                        return "Current Refresh token is invalid or expired ";
                    }

                    settings.UpdateAndSave(oAuth2AccessToken);

                    return "After call to refresh" + Environment.NewLine + settings.ToString();
                }
                return Environment.NewLine + "Ready";
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
        public List<string> FindSpecialSymbol(List<string> listKeyword)
        {
            for (int i = 0; i < listKeyword.Count; i++)
            {
                listKeyword[i] = (listKeyword[i].Replace(" ", "")).Replace(tempSpace, "");
                if (FindCyrillicSymbol(listKeyword[i]))
                {
                    foreach (KeyValuePair<string, string> pair in charReplace)
                    {
                        listKeyword[i] = listKeyword[i].Replace(pair.Key, pair.Value);
                    }
                }
            }
            return listKeyword;
        }
        public async Task FindDescPack(string partNumber)
        {
            try
            {
                var response = await client.KeywordSearch(partNumber);

                subStr = "\"ExactManufacturerProductsCount\":";
                startIndex = response.IndexOf(subStr);
                string tmpResponse = response.Substring(startIndex);
                startIndex = tmpResponse.IndexOf(subStr);
                endIndex = tmpResponse.IndexOf(',');
                int num = int.Parse((tmpResponse.Substring(startIndex + subStr.Length, endIndex - (startIndex + subStr.Length))).Trim(charToTrim));
                if (num != 0)
                    FindFamily(FindPackage(tmpResponse));
                else
                    FindFamily(FindPackage(response));
            }
            catch (Exception)
            {
                throw;
            }
        }
        private void FindFamily(string response)
        {
            subStr = "\"Family\",\"Value\":";
            if (response.IndexOf(subStr) != -1)
            {
                startIndex = response.IndexOf(subStr);
                response = response.Substring(startIndex);
                subStr = "\"Value\":";
                startIndex = response.IndexOf(subStr);
                endIndex = response.IndexOf('}');

                Family.Add((response.Substring(startIndex + subStr.Length, endIndex - (startIndex + subStr.Length))).Trim(charToTrim));
            }
            else
            {
                Family.Add("null");
            }
        }
        private string FindPackage(string response)
        {
            subStr = "\"Parameter\":\"Package / Case\",";
            if (response.IndexOf(subStr) != -1)
            {
                startIndex = response.IndexOf(subStr);
                response = response.Substring(startIndex);
                subStr = "\"Value\":";
                startIndex = response.IndexOf(subStr);
                endIndex = response.IndexOf("}");
                Package.Add((response.Substring(startIndex + subStr.Length, endIndex - (startIndex + subStr.Length))).Trim(charToTrim));
            }
            else
            {
                Package.Add("null");
            }

            return response;
        }
        public void FindPassiveComponents(string pathToDoc, int numSheet, string family)
        {
            try
            {
                ActionWithExcel = new ActionWithExcel();
                bool checkPassiveComponent = ActionWithExcel.UpdateExcelDoc(pathToDoc, numSheet, family);    //Checking for passive components                                                                                    
                if (checkPassiveComponent)
                    passiveComponents.Add("Passive");
                else
                    passiveComponents.Add("null");
            }
            catch(Exception)
            {
                throw;
            }
        }
        public void FindUniversalEquipment(string pathToDoc, int numSheet, string family)
        {
            try
            {
                ActionWithExcel = new ActionWithExcel();
                universalEquipment.Add(ActionWithExcel.UpdateExcelDocForReadUniversalEquipmentFile(pathToDoc, numSheet, family));
            }
            catch(Exception)
            {
                throw;
            }
        }
        public void FindEngineer(string pathToDoc, int numSheet, string family)
        {
            try
            {
                ActionWithExcel = new ActionWithExcel();
                engineer.Add(ActionWithExcel.UpdateExcelDocForReadEngineer(pathToDoc, numSheet, family));
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void FindDifficulty(string pathToDoc, int numSheet, string family)
        {
            try
            {
                ActionWithExcel = new ActionWithExcel();
                difficulty.Add(ActionWithExcel.UpdateExcelDocForReadDifficulty(pathToDoc, numSheet, family));
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void FindMotherBoard(string pathToDoc, int numSheet, string partNumber)
        {
            try
            {
                ActionWithExcel = new ActionWithExcel();
                motherBoard.Add(ActionWithExcel.UpdateExcelDocForReadMotherBoard(pathToDoc, numSheet, partNumber));
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void FindMotherBoardTrim(string pathToDoc, int numSheet, string partNumber)
        {
            try
            {
                ActionWithExcel = new ActionWithExcel();
                motherBoardTrim.Add(ActionWithExcel.UpdateExcelDocForReadMotherBoard(pathToDoc, numSheet, partNumber.TrimEnd(PartNumberTrim)));
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
