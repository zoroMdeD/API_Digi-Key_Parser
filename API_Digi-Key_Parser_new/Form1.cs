﻿using System;
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
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace API_Digi_Key_Parser_new
{
    public partial class Form1 : Form
    {
        string[] MassPathFile;  //Массив путей к файлам

        string PathInfoPartNumbers = string.Empty;
        string PathInfoPartNumberPass = string.Empty;
        string PathInfoEngineers = string.Empty;
        string PathInfoAdapters = string.Empty;
        string PathInfoMotherBoard = string.Empty;

        bool CheckBtnWrkSlt = false;
        bool CheckBtnParsing = false;
        AboutBox1 about_program = new AboutBox1();
        List<string> InputNameSheets = new List<string>();
        List<string> InputDesc = new List<string>();
        string[] MassDescription;
        string[] MassPackage;
        string[] MassTmp;
        string[] allFoundFiles = new string[1];
        Parser Parser;

        public delegate void MyDelegate();      //Для доступа к элементам из другого потока с передачей параметров

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
            if (!CheckBtnParsing)   //Checking tap button "Parsing"
            {
                if (toolStripTextBox1.TextLength > 1)   //System.NullReferenceException
                {
                    //if ((toolStripTextBox4.TextLength > 1) || (teststr.Contains("\\Test_API_DigiKey")))   //System.NullReferenceException
                    if (allFoundFiles[0].Contains("\\Test_API_DigiKey"))
                    {
                        oAuthToolStripMenuItem.Enabled = false;
                        saveToolStripMenuItem.Enabled = false;
                        pathToolStripMenuItem.Enabled = false;
                        CheckBtnParsing = true;
                        label1.Text = "Processing...";
                        TaskRun(PathInfoPartNumbers);
                    }
                    else
                    {
                        DialogResult result;
                        result = MessageBox.Show("Please select the working directory in the settings!", "Working directory not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                }
                else
                {
                    DialogResult result;
                    result = MessageBox.Show("Please select the path to the file!", "File not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string[] MassDesc = new string[] { "Temperature Sensors - Analog and Digital Output"
                                             , "Out of Bounds", "Memory"
                                             , "Memory"
                                             , "Embedded - Microcontrollers"
                                             , "Clock/Timing - Clock Generators, PLLs, Frequency Synthesizers"
                                             , "Memory"
                                             , "DC DC Converters"
                                             , "Memory"
                                             , "Out of Bounds"
                                             , "RF Amplifiers"
                                             , "Out of Bounds"
                                             , "Memory"
                                             , "Embedded - CPLDs (Complex Programmable Logic Devices)"
                                             , "Interface - Drivers, Receivers, Transceivers"
                                             , "Interface - Drivers, Receivers, Transceivers"
                                             , "Interface - Drivers, Receivers, Transceivers"
                                             , "Interface - Drivers, Receivers, Transceivers"
                                             , "Memory"
                                             , "Memory" };
            string[] MassPartNumbers = new string[] { "ADT7411ARQZ"
                                                    , "S29AL032D70TFI000"
                                                    , "AT49BV162A-70TI"
                                                    , "NAND512W3A2BN6E"
                                                    , "AT91R40008-66AU"
                                                    , "CY23EP09ZXI-1H"
                                                    , "CY62157EV30LL-45ZSXI"
                                                    , "DCP021212DU"
                                                    , "FM28V020-SG"
                                                    , "GDPXA255A0E400"
                                                    , "MGA-72543"
                                                    , "MT48LC16M16A2P-6A IT:G"
                                                    , "TE28F256P30B95"
                                                    , "XC9572XL-10VQG64I"
                                                    , "ADM208ARZ"
                                                    , "ADM3075EARZ "
                                                    , "ADM3202ARNZ"
                                                    , "ADM489ARZ"
                                                    , "AT25M02-SSHM-T"
                                                    , "AT29LV256-20JI" };
            string[] MassHead = new string[] { "PartNumber"
                                             , "Description"
                                             , "Package"
                                             , "Engineer"
                                             , "Difficulty"
                                             , "Adapters"
                                             , "MotherBoard"};
            //if (toolStripTextBox1.TextLength > 1)
            //{
                Excel.Application excelApp = new Excel.Application();
                // Создаём экземпляр рабочий книги Excel
                Excel.Workbook workBook;
                // Создаём экземпляр листа Excel
                Excel.Worksheet workSheet;


                workBook = excelApp.Workbooks.Add();
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

                // Заполняем шапку таблицы
                for (int i = 0, j = 1; i < MassHead.Length; i++, j++)
                {
                    workSheet.Cells[1, j] = MassHead[i];
                }
                // Заполняем наименование микросхем (1-й столбец)
                for (int i = 0, j = 2; i < MassPartNumbers.Length; i++, j++)
                {
                    workSheet.Cells[j, 1] = MassPartNumbers[i];
                }
                // Заполняем описание микросхем (2-ой столбец)
                for (int i = 0, j = 2; i < MassDesc.Length; i++, j++)
                {
                    workSheet.Cells[j, 2] = MassDesc[i];
                }

                workSheet.Columns.EntireColumn.AutoFit();
                // Открываем созданный excel-файл
                excelApp.Visible = true;
                excelApp.UserControl = true;
            //}
            //else
            //{
            //    DialogResult result;
            //    result = MessageBox.Show("Please select the path to the file!", "File not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            //}
        }
        private void pathToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!CheckBtnWrkSlt)
            {
                CheckBtnWrkSlt = true;
                GetPathDirectory();
                Thread.Sleep(500);
                toolStripTextBox1.Text = Open_dialog();
                FindAllFileOnLocalServer(toolStripTextBox4.Text);
            }
            else
            {
                DialogResult result;
                result = MessageBox.Show("Please select the working directory in the settings!", "Working directory not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
        string Open_dialog()
        {
            openFileDialog1.Filter = "Excel files (*.xls;*.xlsx;)|*.xls;*.xlsx";
            openFileDialog1.FileName = "Select a some file";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                PathInfoPartNumbers = openFileDialog1.FileName;
                label1.Text = "File selected";
            }
            return PathInfoPartNumbers;
        }
        void FindAllFileOnLocalServer(string strThread)
        {
            Thread Thread_One = new Thread(new ParameterizedThreadStart(Thread_ReadFileLocalServ));                //Создаем новый объект потока (Thread)
            Thread_One.IsBackground = true;                                                           //Поток является фоновым
            Thread_One.Start(strThread);                                                                       //запускаем поток
        }
        void Thread_ReadFileLocalServ(object s)
        {
            string strPath = (string)s;
            try
            {
                int ienum = 0;
                allFoundFiles[0] = @strPath;//Directory.GetDirectories(@strPath);//, "Test_API_DigiKey", SearchOption.AllDirectories);
                BeginInvoke(new MyDelegate(CheckCorrectPath));
                //textBox1.AppendText(allFoundFiles[0] + Environment.NewLine);    //for debug
                RecursiveFileProcessor RecursiveFileProcessor = new RecursiveFileProcessor(allFoundFiles);
                RecursiveFileProcessor.RunProcessor(RecursiveFileProcessor.GetPath);
                MassPathFile = new string[RecursiveFileProcessor.OutPath.Count];
                foreach (string item in RecursiveFileProcessor.OutPath)
                {
                    MassPathFile[ienum] = GetLnkTarget(RecursiveFileProcessor.OutPath[ienum]);
                    ienum++;
                    //textBox1.AppendText(item + Environment.NewLine);    //for debug
                }
                CheckBtnWrkSlt = false;
            }
            catch(Exception)
            { 
                ; 
            }
            //BeginInvoke(new MyDelegate(test2));
        }
        public static string GetLnkTarget(string lnkPath)   //Method get path an link
        {
            var shl = new Shell32.Shell();         // Move this to class scope
            lnkPath = Path.GetFullPath(lnkPath);
            var dir = shl.NameSpace(Path.GetDirectoryName(lnkPath));
            var itm = dir.Items().Item(Path.GetFileName(lnkPath));
            var lnk = (Shell32.ShellLinkObject)itm.GetLink;
            return lnk.Target.Path;
        }
        void CheckCorrectPath()
        {
            if (allFoundFiles.Length > 0)
                textBox1.AppendText(allFoundFiles[0] + Environment.NewLine);    //for debug
            else
            {
                DialogResult result;
                result = MessageBox.Show("The working directory was not found in the selected directory!", "Invalid directory", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
        string FindPathToFile(string name)
        {
            for (int i = 0; i < MassPathFile.Length; i++)
            {
                if (MassPathFile[i].Contains(name))
                    return MassPathFile[i];
            }
            return "null";
        }
        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void parsingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!CheckBtnParsing)   //Checking tap button "Parsing"
            {
                if (toolStripTextBox1.TextLength > 1)
                {
                    if (toolStripTextBox4.TextLength > 1)   //System.NullReferenceException
                    {
                        oAuthToolStripMenuItem.Enabled = false;
                        saveToolStripMenuItem.Enabled = false;
                        pathToolStripMenuItem.Enabled = false;
                        CheckBtnParsing = true;
                        label1.Text = "Processing...";
                        TaskRun(PathInfoPartNumbers);
                    }
                    else
                    {
                        DialogResult result;
                        result = MessageBox.Show("Please select the working directory in the settings!", "Working directory not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                }
                else
                {
                    DialogResult result;
                    result = MessageBox.Show("Please select the path to the file!", "File not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
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
        void GetPathDirectory()
        {
            openFileDialog1.ValidateNames = false;
            openFileDialog1.CheckFileExists = false;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.FileName = "Select a some directory";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                toolStripTextBox4.Text = Path.GetDirectoryName(openFileDialog1.FileName);
            }
        }
        async void TaskRun(string Path)
        {
            try
            {
                Parser = new Parser(FindPathToFile(@"\InfoPartNumberPass.xlsx"));     //Путь для файла должен быть динамическим!!!
                Task task = Parser.ParserInit();
                await task;

                ActionWithExcel ActionWithExcel = new ActionWithExcel();
                InputDesc = ActionWithExcel.UpdateExcelDoc(Path, 0);

                MassTmp = new string[InputDesc.Count];
                MassDescription = new string[InputDesc.Count];
                MassPackage = new string[InputDesc.Count];

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
                MassTmp[i] = await Parser.FindDescriprions(InputDesc[i]);
                if (MassTmp[i].IndexOf('#') > 0)
                {
                    MassDescription[i] = MassTmp[i].Substring(0, MassTmp[i].IndexOf('#'));
                    MassPackage[i] = MassTmp[i].Substring(MassTmp[i].IndexOf('#'));
                    textBox1.AppendText(MassDescription[i] + Environment.NewLine);  //for debug
                }
                else
                {
                    MassDescription[i] = MassTmp[i];
                    MassPackage[i] = "null";
                    textBox1.AppendText(MassDescription[i] + Environment.NewLine);  //for debug
                }
                p.Report(i);
            }
            label1.Text = "Completed";
            CheckBtnParsing = false;
            oAuthToolStripMenuItem.Enabled = true;
            saveToolStripMenuItem.Enabled = true;
            pathToolStripMenuItem.Enabled = true;
        }
    }
}
