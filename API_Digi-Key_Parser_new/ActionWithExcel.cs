using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace API_Digi_Key_Parser_new
{
    public class ActionWithExcel
    {
        private List<string> listNameSheets;
        private ConnectToExcel ConnectToExcel;
        private ListOfPartNumbers ListOfPartNumbers;
        private List<string> outMassInfo;
        private bool outFlagPass;

        public List<string> ListNameSheets
        {
            get
            {
                return listNameSheets;
            }
            private set
            {
                listNameSheets = value;
            }
        }
        public List<string> OutMassInfo
        {
            get
            {
                return outMassInfo;
            }
            private set
            {
                outMassInfo = value;
            }
        }
        public bool OutFlagPass
        {
            get
            {
                return outFlagPass;
            }
            private set
            {
                outFlagPass = value;
            }
        }

        public ActionWithExcel()
        {

        }
        public List<string> UpdateExcelDoc(string Path, int NumSheet)
        {
            ConnectToExcel = new ConnectToExcel(@Path);
            listNameSheets = ConnectToExcel.GetWorksheetNames(ConnectToExcel);
            ListOfPartNumbers = new ListOfPartNumbers(@Path, listNameSheets[NumSheet]);
            outMassInfo = ListOfPartNumbers.GetListInfoExcelDoc(ConnectToExcel);
            return outMassInfo;
        }
        public bool UpdateExcelDoc(string Path, int NumSheet, string Family)
        {
            ConnectToExcel = new ConnectToExcel(@Path);
            listNameSheets = ConnectToExcel.GetWorksheetNames(ConnectToExcel);
            ListOfPartNumbers = new ListOfPartNumbers(@Path, listNameSheets[NumSheet]);
            outFlagPass = ListOfPartNumbers.GetListInfoExcelDoc(ConnectToExcel, Family);
            return outFlagPass;
        }
    }
}
