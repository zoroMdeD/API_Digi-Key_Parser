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
        string Path = string.Empty;
        AboutBox1 about_program = new AboutBox1();
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
            label1.Text = "No file selected";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (toolStripTextBox1.TextLength > 1)   //System.NullReferenceException
            {
                label1.Text = "Processing...";
                TaskRun(Path);
            }
            else
            {
                DialogResult result;
                result = MessageBox.Show("Please select the path to the file!", "File not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (toolStripTextBox1.TextLength > 1)
            {

            }
            else
            {
                DialogResult result;
                result = MessageBox.Show("Please select the path to the file!", "File not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
        private void pathToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripTextBox1.Text = Open_dialog();
        }
        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void parsingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (toolStripTextBox1.TextLength > 1)
            {

            }
            else
            {
                DialogResult result;
                result = MessageBox.Show("Please select the path to the file!", "File not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void oAuthToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (toolStripTextBox1.TextLength > 1)
            {

            }
            else
            {
                DialogResult result;
                result = MessageBox.Show("Please select the path to the file!", "File not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (toolStripTextBox1.TextLength > 1)
            {

            }
            else
            {
                DialogResult result;
                result = MessageBox.Show("Please select the path to the file!", "File not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
        private void viewHelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            about_program.ShowDialog();
        }

        async void TaskRun(string Path)
        {
            try
            {
                ConnectToExcel = new ConnectToExcel(@Path);
                ConxObject = new ConnectToExcel(@Path);
                InputNameSheets = ConnectToExcel.GetWorksheetNames(ConxObject);
                ListOfPartNumbers ListOfPartNumbers = new ListOfPartNumbers(@Path, InputNameSheets[0]);
                InputDesc = ListOfPartNumbers.GetListOfPartNumbers(ConxObject);
                MassDescription = new string[InputDesc.Count];

                Parser = new Parser();
                Task task = Parser.ParserInit();
                await task;

                // since this is a UI event, instantiating the Progress class
                // here will capture the UI thread context
                var progress = new Progress<int>(i => progressBar1.Value = i);
                progressBar1.Minimum = 0;
                progressBar1.Maximum = InputDesc.Count - 1;
                // pass this instance to the background task
                _ = OutData(progress);
            }
            catch(Exception e)
            {
                textBox1.AppendText(e.Message);
            }
        }

        async Task OutData(IProgress<int> p)
        {
            for (int i = 0; i < InputDesc.Count; i++)
            {
                MassDescription[i] = await Parser.FindDescriprions(InputDesc[i]);
                textBox1.AppendText(MassDescription[i] + Environment.NewLine);
                p.Report(i);
            }
            label1.Text = "Completed";
        }
        string Open_dialog()
        {
            openFileDialog1.Filter = "Excel files (*.xls;*.xlsx;)|*.xls;*.xlsx";
            openFileDialog1.FileName = "Some File";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Path = openFileDialog1.FileName;
                label1.Text = "File selected";
            }
            return Path;
        }

    }
}
