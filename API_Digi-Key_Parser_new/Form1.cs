using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ApiClient.Models;
using ApiClient.OAuth2;
using ApiClient;
using Common.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Threading;
using System.ComponentModel;

namespace API_Digi_Key_Parser_new
{
    public partial class Form1 : Form
    {
        List<string> InputNameSheets = new List<string>();
        List<string> InputDesc = new List<string>();
        string[] MassDescription;
        ConnectToExcel ConnectToExcel;
        ConnectToExcel ConxObject;
        Parser Parser;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            ConnectToExcel = new ConnectToExcel(@"D:\TestExcel.xlsx");
            ConxObject = new ConnectToExcel(@"D:\TestExcel.xlsx");
            InputNameSheets = ConnectToExcel.GetWorksheetNames(ConxObject);
            ListOfPartNumbers ListOfPartNumbers = new ListOfPartNumbers(@"D:\TestExcel.xlsx", InputNameSheets[0]);
            InputDesc = ListOfPartNumbers.GetListOfPartNumbers(ConxObject);
            MassDescription = new string[InputDesc.Count];

            Parser = new Parser();
            Task task = Parser.ParserInit();
            await task;

            OutData();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        async void OutData()
        {
            for (int i = 0; i < InputDesc.Count; i++)
            {
                MassDescription[i] = await Parser.FindDescriprions(InputDesc[i]);
            }
        }
    }
}
