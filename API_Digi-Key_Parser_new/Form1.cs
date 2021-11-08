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
        ConnectToExcel ConnectToExcel;
        ConnectToExcel ConxObject;

        static BackgroundWorker EndOperation;
        public delegate void MyDelegate();      //Для доступа к элементам из другого потока с передачей параметров

        public Form1()
        {
            InitializeComponent();

            EndOperation = new BackgroundWorker();
            EndOperation.WorkerReportsProgress = true;
            EndOperation.DoWork += EndOperation_DoWork;
            EndOperation.ProgressChanged += EndOperation_ProgressChanged;
            EndOperation.RunWorkerCompleted += EndOperation_RunWorkerCompleted;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //List<string> InputNameSheets = new List<string>();
            //List<string> InputDesc = new List<string>();
            //
            //ConnectToExcel ConnectToExcel = new ConnectToExcel(@"D:\TestExcel.xlsx");
            //ConnectToExcel ConxObject = new ConnectToExcel(@"D:\TestExcel.xlsx");
            //InputNameSheets = ConnectToExcel.GetWorksheetNames(ConxObject);
            //ListOfPartNumbers ListOfPartNumbers = new ListOfPartNumbers(@"D:\TestExcel.xlsx", InputNameSheets[0]);
            //InputDesc = ListOfPartNumbers.GetListOfPartNumbers(ConxObject);
            //Parser Parser = new Parser();
            //Task task = Parser.ParserInit();
            //await task;
            //
            //progressBar1.Minimum = 0;
            //progressBar1.Maximum = InputDesc.Count;
            //Progress2 = InputDesc.Count;
            //EndOperation.RunWorkerAsync(null);
            //
            //for (int i = 0; i < InputDesc.Count; i++)
            //{
            //    textBox1.AppendText(await Parser.FindPartNumbers(InputDesc[i]) + Environment.NewLine);
            //}

            ConnectToExcel = new ConnectToExcel(@"D:\TestExcel.xlsx");
            ConxObject = new ConnectToExcel(@"D:\TestExcel.xlsx");
            InputNameSheets = ConnectToExcel.GetWorksheetNames(ConxObject);
            ListOfPartNumbers ListOfPartNumbers = new ListOfPartNumbers(@"D:\TestExcel.xlsx", InputNameSheets[0]);
            InputDesc = ListOfPartNumbers.GetListOfPartNumbers(ConxObject);

            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = ListOfPartNumbers.MassPartNumber.Count - 1;

            EndOperation.RunWorkerAsync(null);
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
         
        async void EndOperation_DoWork(object sender, DoWorkEventArgs e)
        {
            int Progress = 0;
            BackgroundWorker worker = sender as BackgroundWorker;

            //List<string> InputNameSheets = new List<string>();
            //List<string> InputDesc = new List<string>();
            
            //ConnectToExcel ConnectToExcel = new ConnectToExcel(@"D:\TestExcel.xlsx");
            //ConnectToExcel ConxObject = new ConnectToExcel(@"D:\TestExcel.xlsx");
            // = ConnectToExcel.GetWorksheetNames(ConxObject);
            //ListOfPartNumbers ListOfPartNumbers = new ListOfPartNumbers(@"D:\TestExcel.xlsx", InputNameSheets[0]);
            //InputDesc = ListOfPartNumbers.GetListOfPartNumbers(ConxObject);

            Parser Parser = new Parser();
            Task task = Parser.ParserInit();
            await task;
            
            for (int i = 0; i < InputDesc.Count; i++)
            {
                //textBox1.AppendText(await Parser.FindPartNumbers(InputDesc[i]) + Environment.NewLine);
                EndOperation.ReportProgress(Progress++);
            }

            e.Result = "Complete...";
        }

        void EndOperation_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
                textBox1.Text += "Interrupted by user" + Environment.NewLine;
            else if (e.Error != null)
                textBox1.Text += "Interrupted" + Environment.NewLine;
            else
            {
                textBox1.Text += e.Result + Environment.NewLine;
            }
            //progressBar1.Value = 0;
        }

        void EndOperation_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }
    }
}
