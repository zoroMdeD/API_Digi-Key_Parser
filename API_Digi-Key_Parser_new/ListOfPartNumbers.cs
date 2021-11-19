using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace API_Digi_Key_Parser_new
{
    class ListOfPartNumbers
    {
        private string pathToExcelFile = string.Empty;
        private string nameOfSheet = string.Empty;
        private List<string> massPartNumber;

        public string PathToExcelFile
        {
            get
            {
                return pathToExcelFile;
            }
            private set
            {
                pathToExcelFile = value;
            }
        }
        public string NameOfSheet
        {
            get
            {
                return nameOfSheet;
            }
            private set
            {
                nameOfSheet = value;
            }
        }
        public List<string> MassPartNumber
        {
            get
            {
                return massPartNumber;
            }
            private set
            {
                massPartNumber = value;
            }
        }

        public string PartNumber { get; set; }          //Input file PartNumbers
        public string PartNumberPass { get; set; }      //File passive components
        public string Description { get; set; }         //Universal equipment
        public string BuildNumber { get; set; }         //Universal equipment
        public string Engineer { get; set; }            //Input file Engineers
        public int Difficulty { get; set; }          //Input file Engineers

        public ListOfPartNumbers(string PathToExcelFile, string NameOfSheet = "Лист1")
        {
            this.PathToExcelFile = PathToExcelFile;
            this.NameOfSheet = NameOfSheet;
        }
        public ListOfPartNumbers()
        {

        }

        public List<string> GetListInfoExcelDoc(ConnectToExcel ConnectToExcel)
        {
            MassPartNumber = new List<string>();
            //Query a worksheet with a header row (sintax SQL-Like LINQ)
            var GetSheet = from a in ConnectToExcel.UrlConnexion.Worksheet<ListOfPartNumbers>(nameOfSheet)
                           select a;
            foreach (var result in GetSheet)
            {
                MassPartNumber.Add(result.PartNumber);
            }
            return MassPartNumber;
        }
        //Method checking for passive components
        public bool GetListInfoExcelDoc(ConnectToExcel ConnectToExcel, string Family)
        {
            bool match = false;
            //Query a worksheet with a header row (sintax SQL-Like LINQ)
            var GetSheet = from a in ConnectToExcel.UrlConnexion.Worksheet<ListOfPartNumbers>(nameOfSheet)
                           select a;
            foreach (var result in GetSheet)
            {
                if(Family == result.PartNumberPass)
                {
                    match = true;
                    break;
                }
            }
            return match;
        }
        public string GetListInfoExcelDocUniversalEquipment(ConnectToExcel ConnectToExcel, string Family)
        {
            string BuildNumber = string.Empty;
            //Query a worksheet with a header row (sintax SQL-Like LINQ)
            var GetSheet = from a in ConnectToExcel.UrlConnexion.Worksheet<ListOfPartNumbers>(nameOfSheet)
                           select a;
            foreach (var result in GetSheet)
            {
                if (Family == result.Description)
                {
                    BuildNumber = result.BuildNumber;
                    return BuildNumber;
                }
            }
            return "null";
        }
        public string GetListInfoExcelDocEngineer(ConnectToExcel ConnectToExcel, string Family)
        {
            string Engineer = string.Empty;
            //Query a worksheet with a header row (sintax SQL-Like LINQ)
            var GetSheet = from a in ConnectToExcel.UrlConnexion.Worksheet<ListOfPartNumbers>(nameOfSheet)
                           select a;
            foreach (var result in GetSheet)
            {
                if (Family == result.Description)
                {
                    Engineer = result.Engineer;
                    return Engineer;
                }
            }
            return "null";
        }
        public int GetListInfoExcelDocDifficulty(ConnectToExcel ConnectToExcel, string Family)
        {
            int Difficulty = 0;
            //Query a worksheet with a header row (sintax SQL-Like LINQ)
            var GetSheet = from a in ConnectToExcel.UrlConnexion.Worksheet<ListOfPartNumbers>(nameOfSheet)
                           select a;
            foreach (var result in GetSheet)
            {
                if (Family == result.Description)
                {
                    Difficulty = result.Difficulty;
                    return Difficulty;
                }
            }
            return -1;
        }
    }
}
