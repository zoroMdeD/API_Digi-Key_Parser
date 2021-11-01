using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace API_Digi_Key_Parser_new
{
    class ListOfPartNumbers
    {
        private string pathToExcelFile = "none";
        private string nameOfSheet = "none";

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
        public string PartNumber
        {
            get;
            set;
        }

        public ListOfPartNumbers(string PathToExcelFile, string NameOfSheet = "Лист1")
        {
            this.PathToExcelFile = PathToExcelFile;
            this.NameOfSheet = NameOfSheet;
        }
        public ListOfPartNumbers()
        {

        }
        public List<string> GetListOfPartNumbers()
        {
            List<string> MassPartNumber = new List<string>();
            ConnectToExcel ConxObject = new ConnectToExcel(pathToExcelFile);
            //Query a worksheet with a header row (sintax SQL-Like LINQ)
            var GetSheet = from a in ConxObject.UrlConnexion.Worksheet<ListOfPartNumbers>(nameOfSheet)
                           select a;
            foreach (var result in GetSheet)
            {
                MassPartNumber.Add(result.PartNumber);
            }
            return MassPartNumber;
        }
    }
}
