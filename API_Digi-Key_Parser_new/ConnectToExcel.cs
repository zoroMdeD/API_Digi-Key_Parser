using System;
using LinqToExcel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace API_Digi_Key_Parser_new
{
    public class ConnectToExcel
    {
        public string _pathExcelFile;
        public ExcelQueryFactory _urlConnexion;
        private List<string> nameWorksheet;

        public List<string> NameWorksheet
        {
            get
            {
                return nameWorksheet;
            }
            private set
            {
                nameWorksheet = value;
            }
        }
        public ConnectToExcel(string path)
        {
            this._pathExcelFile = path;
            this._urlConnexion = new ExcelQueryFactory(_pathExcelFile);
        }
        public string PathExcelFile
        {
            get
            {
                return _pathExcelFile;
            }
        }
        public ExcelQueryFactory UrlConnexion
        {
            get
            {
                return _urlConnexion;
            }
        }

        public static List<string> GetWorksheetNames(ConnectToExcel ConxObject)
        {
            try
            {
                List<string> MassWorksheetNames = new List<string>();
                var worksheetNames = ConxObject.UrlConnexion.GetWorksheetNames();
                foreach (var result in worksheetNames)
                {
                    MassWorksheetNames.Add(result);
                }

                return MassWorksheetNames;
            }
            catch(Exception)
            {
                throw;
            }
        }
    }
}
