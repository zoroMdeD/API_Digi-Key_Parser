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

        public string PartNumber    //Исходники
        {
            get;
            set;
        }
        public string PartNumberPass    //Проверка на пассивку
        {
            get;
            set;
        }
        //Для поиска адаптеров
        //public string BuildNumber    
        //{
        //    get;
        //    set;
        //}

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
    }
}
