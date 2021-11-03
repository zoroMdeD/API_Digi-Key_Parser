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

namespace API_Digi_Key_Parser_new
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            ConnectToExcel ConnectToExcel = new ConnectToExcel(@"D:\TestExcel.xlsx");
            ListOfPartNumbers ListOfPartNumbers = new ListOfPartNumbers(@"D:\TestExcel.xlsx");
            List<string> InputData = new List<string>();
            InputData = ListOfPartNumbers.GetListOfPartNumbers();
            Parser Parser = new Parser();
            Task task = Parser.ParserInit();
            await task;

            for (int i = 0; i < InputData.Count; i++)
            {
                textBox1.AppendText(await Parser.FindPartNumbers(InputData[i]) + Environment.NewLine);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
