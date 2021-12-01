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
            try
            {
                ConnectToExcel = new ConnectToExcel(@Path);
                listNameSheets = ConnectToExcel.GetWorksheetNames(ConnectToExcel);
                ListOfPartNumbers = new ListOfPartNumbers(@Path, listNameSheets[NumSheet]);
                outMassInfo = ListOfPartNumbers.GetListInfoExcelDoc(ConnectToExcel);
                return outMassInfo;
            }
            catch (Exception) 
            { 
                throw; 
            }
        }
        public bool UpdateExcelDoc(string Path, int NumSheet, string Family)
        {
            try
            {
                ConnectToExcel = new ConnectToExcel(@Path);
                listNameSheets = ConnectToExcel.GetWorksheetNames(ConnectToExcel);
                ListOfPartNumbers = new ListOfPartNumbers(@Path, listNameSheets[NumSheet]);
                outFlagPass = ListOfPartNumbers.GetListInfoExcelDoc(ConnectToExcel, Family);
                return outFlagPass;
            }
            catch (Exception) 
            { 
                throw; 
            }
        }
        public string UpdateExcelDocForReadUniversalEquipmentFile(string Path, int NumSheet, string Family)
        {
            try
            {
                string outString = string.Empty;
                ConnectToExcel = new ConnectToExcel(@Path);
                listNameSheets = ConnectToExcel.GetWorksheetNames(ConnectToExcel);
                ListOfPartNumbers = new ListOfPartNumbers(@Path, listNameSheets[NumSheet]);
                outString = ListOfPartNumbers.GetListInfoExcelDocUniversalEquipment(ConnectToExcel, Family);
                return outString;
            }
            catch (Exception) 
            { 
                throw; 
            }
        }
        public string UpdateExcelDocForReadEngineer(string Path, int NumSheet, string Family)
        {
            try
            {
                string outString = string.Empty;
                ConnectToExcel = new ConnectToExcel(@Path);
                listNameSheets = ConnectToExcel.GetWorksheetNames(ConnectToExcel);
                ListOfPartNumbers = new ListOfPartNumbers(@Path, listNameSheets[NumSheet]);
                outString = ListOfPartNumbers.GetListInfoExcelDocEngineer(ConnectToExcel, Family);
                return outString;
            }
            catch (Exception) 
            { 
                throw; 
            }
        }
        public int UpdateExcelDocForReadDifficulty(string Path, int NumSheet, string Family)
        {
            try
            {
                int outValue = 0;
                ConnectToExcel = new ConnectToExcel(@Path);
                listNameSheets = ConnectToExcel.GetWorksheetNames(ConnectToExcel);
                ListOfPartNumbers = new ListOfPartNumbers(@Path, listNameSheets[NumSheet]);
                outValue = ListOfPartNumbers.GetListInfoExcelDocDifficulty(ConnectToExcel, Family);
                return outValue;
            }
            catch (Exception) 
            { 
                throw; 
            }
        }
        public string UpdateExcelDocForReadMotherBoard(string Path, int NumSheet, string PartNumber)
        {
            try
            {
                string outString = string.Empty;
                ConnectToExcel = new ConnectToExcel(@Path);
                listNameSheets = ConnectToExcel.GetWorksheetNames(ConnectToExcel);
                ListOfPartNumbers = new ListOfPartNumbers(@Path, listNameSheets[NumSheet]);
                outString = ListOfPartNumbers.GetListInfoExcelDocMotherBoard(ConnectToExcel, PartNumber);
                return outString;
            }
            catch (Exception) 
            { 
                throw; 
            }
        }
    }
}
