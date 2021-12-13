using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ApiClient.Models;
using ApiClient.OAuth2;
using ApiClient;
using ApiClient.Exception;
using Common.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Threading;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Web.Http;

namespace API_Digi_Key_Parser_new
{
    public partial class Form1 : Form
    {
        string[] MassPathFile;  //Массив путей к файлам

        string PathInfoPartNumbers = string.Empty;

        bool CheckBtnWrkSlt = false;
        bool CheckBtnParsing = false;
        bool CheckBtnSave = false;
        AboutBox1 about_program = new AboutBox1();
        List<string> InputNameSheets = new List<string>();
        List<string> ProcessedPartNumbers = new List<string>();
        string[] WorkDirPath = new string[1];
        Parser Parser;

        public delegate void MyDelegate();      //Для доступа к элементам из другого потока с передачей параметров
        static BackgroundWorker DocBuild;

        CancellationTokenSource cts;

        public Form1()
        {
            InitializeComponent();

            yesToolStripMenuItem.Enabled = false;
            toolStripTextBox4.Text = @"X:\DataBase\Parser\Parser_API";    //X:\DataBase\Parser\Parser_API     X:\DataBase\Test_2\Test_API_DigiKey

            DocBuild = new BackgroundWorker();
            DocBuild.WorkerSupportsCancellation = true;
            DocBuild.WorkerReportsProgress = true;
            DocBuild.DoWork += DocBuild_DoWork;
            DocBuild.ProgressChanged += DocBuild_ProgressChanged;
            DocBuild.RunWorkerCompleted += DocBuild_RunWorkerCompleted;
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
                    if (WorkDirPath[0].Contains("\\Parser_API"))    //"\\Parser_API"      "\\Test_API_DigiKey"
                    {
                        oAuthToolStripMenuItem.Enabled = false;
                        saveToolStripMenuItem.Enabled = false;
                        pathToolStripMenuItem.Enabled = false;
                        CheckBtnParsing = true;
                        CheckBtnSave = false;
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
            if (CheckBtnSave)
            {
                SaveExcelDoc();
            }
            else if(!CheckBtnParsing)
            {
                DialogResult result;
                result = MessageBox.Show("Please run the parser!", "Nothing to save", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else if(CheckBtnParsing)
            {
                DialogResult result;
                result = MessageBox.Show("Please wait for the parser to finish working!", "Nothing to save", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
        private void pathToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (!CheckBtnWrkSlt)
                {
                    CheckBtnWrkSlt = true;
                    if (yesToolStripMenuItem.Enabled != false)
                        GetPathDirectory();
                    Thread.Sleep(250);
                    toolStripTextBox1.Text = Open_dialog();
                    FindAllFileOnLocalServer(toolStripTextBox4.Text);
                }
                else
                {
                    DialogResult result;
                    result = MessageBox.Show("Please select the working directory in the settings!", "Working directory not selected", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
            catch(Exception ex)
            {
                textBox1.AppendText(Environment.NewLine + "Directory path error: " + ex.Message);
                label1.Text = "Error";
                CheckBtnParsing = false;
                oAuthToolStripMenuItem.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                pathToolStripMenuItem.Enabled = true;
            }
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
                        CheckBtnSave = false;
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
        private async void oAuthToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                oAuthToolStripMenuItem.Enabled = false;

                OAuth OAuth = new OAuth();
                string StrOut = await OAuth.Authorize();

                textBox1.AppendText(Environment.NewLine + StrOut);    //for info
            }
            catch (Exception ex)
            {
                textBox1.AppendText(Environment.NewLine + "Authorization error: " + ex.Message);
                label1.Text = "Error";
                CheckBtnParsing = false;
                oAuthToolStripMenuItem.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                pathToolStripMenuItem.Enabled = true;
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (CheckBtnSave)
            {
                SaveExcelDoc();
            }
            else if (!CheckBtnParsing)
            {
                DialogResult result;
                result = MessageBox.Show("Please run the parser!", "Nothing to save", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else if (CheckBtnParsing)
            {
                DialogResult result;
                result = MessageBox.Show("Please wait for the parser to finish working!", "Nothing to save", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
        private void viewHelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            about_program.ShowDialog();
        }
        private void yesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            noToolStripMenuItem.Enabled = true;
            yesToolStripMenuItem.Enabled = false;
        }
        private void noToolStripMenuItem_Click(object sender, EventArgs e)
        {
            yesToolStripMenuItem.Enabled = true;
            noToolStripMenuItem.Enabled = false;
        }
        private void abortToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (CheckBtnParsing)
                if (DocBuild.IsBusy)
                    DocBuild.CancelAsync();
                else
                    cts.Cancel();
        }
        void SaveExcelDoc()
        {
            try
            {
                string[] MassHead = new string[] { "PartNumber", "Description", "Package", "Adapters", "MotherBoard", "Engineer", "Difficulty" };

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workBook;                // Создаём экземпляр рабочий книги Excel
                Excel.Worksheet workSheet;              // Создаём экземпляр листа Excel

                workBook = excelApp.Workbooks.Add();
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);


                for (int i = 0, j = 1; i < MassHead.Length; i++, j++)   //Заполняем шапку таблицы
                {
                    workSheet.Cells[1, j] = MassHead[i];
                }
                for (int i = 0, j = 2; i < ProcessedPartNumbers.Count; i++, j++)   //Заполняем таблицу
                {
                    workSheet.Cells[j, 1] = ProcessedPartNumbers[i];
                    //if (Parser.Family[i] != "Out of Bounds")
                        workSheet.Cells[j, 2] = Parser.Family[i];
                    //else
                    //    workSheet.Cells[j, 2] = "null";
                    
                    workSheet.Cells[j, 3] = Parser.Package[i];
                    if (Parser.PassiveComponents[i] != "Passive")
                        workSheet.Cells[j, 4] = "null"; 
                    else
                        workSheet.Cells[j, 4] = Parser.PassiveComponents[i];
                    workSheet.Cells[j, 6] = Parser.Enginner[i];
                    if (Parser.PassiveComponents[i] != "Passive")
                        workSheet.Cells[j, 5] = Parser.MotherBoard[i];
                    if ((Parser.Difficulty[i] == 4) && (Parser.MotherBoard[i] != "null") && (Parser.PassiveComponents[i] != "Passive"))
                        workSheet.Cells[j, 7] = Parser.Difficulty[i] - 1;
                    else
                        workSheet.Cells[j, 7] = Parser.Difficulty[i];
                    if (Parser.MotherBoard[i] == "null")
                        if (Parser.MotherBoardTrim[i] != "null")
                            workSheet.Cells[j, 5] = "match";
                        else if (Parser.PassiveComponents[i] != "Passive")
                            workSheet.Cells[j, 5] = Parser.MotherBoardTrim[i];
                        else
                            workSheet.Cells[j, 5] = Parser.PassiveComponents[i];
                    if ((Parser.PassiveComponents[i] != "Passive") && (Parser.MotherBoard[i] == "null") && (Parser.MotherBoardTrim[i] != "match"))
                        workSheet.Cells[j, 5] = Parser.UniversalEquipment[i];
                }
                workSheet.Columns.EntireColumn.AutoFit();

                /* Enable filter on sheet
                 * Excel.Range target = workSheet.get_Range("A1:G1");
                 * workSheet.Cells.AutoFilter(1, target, Excel.XlAutoFilterOperator.xlAnd, true);
                 */

                excelApp.Visible = true;                    // Открываем созданный excel-файл
                excelApp.UserControl = true;
            }
            catch (Exception ex)
            {
                textBox1.AppendText(Environment.NewLine + "Creating a file error: " + ex.Message);
                label1.Text = "Error";
                CheckBtnParsing = false;
                oAuthToolStripMenuItem.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                pathToolStripMenuItem.Enabled = true;
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
                WorkDirPath[0] = @strPath;
                RecursiveFileProcessor RecursiveFileProcessor = new RecursiveFileProcessor(WorkDirPath);
                RecursiveFileProcessor.RunProcessor(RecursiveFileProcessor.GetPath);
                MassPathFile = new string[RecursiveFileProcessor.OutPath.Count];
                foreach (string item in RecursiveFileProcessor.OutPath)
                {
                    MassPathFile[ienum] = GetLnkTarget(RecursiveFileProcessor.OutPath[ienum]);
                    ienum++;
                }
                CheckBtnWrkSlt = false;
            }
            catch(Exception ex)
            { 
                textBox1.AppendText(Environment.NewLine + "File Detection error: " + ex.Message);
                label1.Text = "Error";
                CheckBtnParsing = false;
                oAuthToolStripMenuItem.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                pathToolStripMenuItem.Enabled = true;
            }
            BeginInvoke(new MyDelegate(CheckCorrectPath));
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
            if (WorkDirPath.Length > 0)
            {
                textBox1.AppendText($"Working directory: {WorkDirPath[0]}" + Environment.NewLine + $"Source file: {PathInfoPartNumbers}");  //for more info
            }
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
                Parser = new Parser();
                string StrOut = await Parser.ParserInit();

                cts = new CancellationTokenSource();
                CancellationToken token = cts.Token;

                textBox1.AppendText(Environment.NewLine + StrOut);    //for info

                ActionWithExcel ActionWithExcel = new ActionWithExcel();
                ProcessedPartNumbers = Parser.FindSpecialSymbol(ActionWithExcel.UpdateExcelDoc(Path, 0));

                // since this is a UI event, instantiating the Progress class
                // here will capture the UI thread context
                var progress = new Progress<int>(i => progressBar1.Value = i);
                progressBar1.Minimum = 0;
                progressBar1.Maximum = (ProcessedPartNumbers.Count - 1)*2;
                // pass this instance to the background task
                _ = OutData(progress, token);
            }
            catch(Exception ex)
            {
                textBox1.AppendText(Environment.NewLine + "Initialization error: " + ex.Message);
                label1.Text = "Error";
                CheckBtnParsing = false;
                oAuthToolStripMenuItem.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                pathToolStripMenuItem.Enabled = true;
            }
        }
        async Task OutData(IProgress<int> p, CancellationToken token)
        {
            bool ExFlag = false;
            string status = "Ready";
            textBox1.AppendText(Environment.NewLine + "Parsing...");
            for (int i = 0; i < ProcessedPartNumbers.Count; i++)
            {
                if (token.IsCancellationRequested)
                {
                    if (!ExFlag)
                        textBox1.AppendText(Environment.NewLine + "Interrupted: The process was interrupted by the user.");    
                    label1.Text = status;
                    p.Report(0);
                    CheckBtnParsing = false;
                    oAuthToolStripMenuItem.Enabled = true;
                    saveToolStripMenuItem.Enabled = true;
                    pathToolStripMenuItem.Enabled = true;
                    cts.Dispose();
                    return;
                }
                try
                {
                    await Parser.FindDescPack(ProcessedPartNumbers[i]);
                }
                catch(AggregateException)
                {
                    ExFlag = true;
                    textBox1.AppendText(Environment.NewLine + "Request limit exceeded: 0 out of 1000 requests left");
                    cts.Cancel();
                    status = "Error";
                }
                catch(NullReferenceException ex)
                {
                    ExFlag = true;
                    textBox1.AppendText(Environment.NewLine + "Authorization error: The token is outdated" + ex.Message);
                    cts.Cancel();
                    status = "Error";
                }
                catch(Exception ex)
                {
                    ExFlag = true;
                    textBox1.AppendText(Environment.NewLine + "Unhandled exception: Something went wrong" + ex.Message);
                    cts.Cancel();
                    status = "Error";
                }
                p.Report(i);

            }
                        /*            catch(HttpResponseException ex)
            {
                textBox1.AppendText(Environment.NewLine + "Request limit exceeded: " + ex.Message);
                label1.Text = "Error";
                CheckBtnParsing = false;
                oAuthToolStripMenuItem.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                pathToolStripMenuItem.Enabled = true;
            }*/
            DocBuild.RunWorkerAsync(null);
        }
        void DocBuild_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                int Progress = ProcessedPartNumbers.Count - 1;
                for (int i = 0; i < ProcessedPartNumbers.Count; i++)
                {
                    if(DocBuild.CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }
                    Parser.FindPassiveComponents(FindPathToFile(@"\InfoPartNumberPass.xlsx"), 0, Parser.Family[i]);
                    Parser.FindUniversalEquipment(FindPathToFile(@"\Universal.xlsx"), 0, Parser.Family[i]);
                    Parser.FindEngineer(FindPathToFile(@"\engineers.xlsx"), 0, Parser.Family[i]);                   //@"\engineers.xlsx"        @"\InfoEngineers.xlsx"
                    Parser.FindDifficulty(FindPathToFile(@"\engineers.xlsx"), 0, Parser.Family[i]);                 //@"\engineers.xlsx"        @"\InfoEngineers.xlsx"
                    Parser.FindMotherBoard(FindPathToFile(@"\Сборки,платы.xlsx"), 8, ProcessedPartNumbers[i]);       //@"\Сборки,платы.xlsx"     @"\InfoMotherBoard.xlsx"
                    Parser.FindMotherBoardTrim(FindPathToFile(@"\Сборки,платы.xlsx"), 8, ProcessedPartNumbers[i]);   //@"\Сборки,платы.xlsx"     @"\InfoMotherBoard.xlsx"
                    DocBuild.ReportProgress(Progress++);
                }
                e.Result = "Completed";
            }
            catch(Exception)
            {
                throw;
            }
        }
        void DocBuild_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                textBox1.AppendText(Environment.NewLine + "Interrupted: The process was interrupted by the user." + Environment.NewLine);
                BeginInvoke(new MyDelegate(RefreshPB));
                label1.Text = "Ready";
                CheckBtnParsing = false;
                oAuthToolStripMenuItem.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                pathToolStripMenuItem.Enabled = true;
            }
            else if (e.Error != null)
            {
                textBox1.AppendText(Environment.NewLine + "Interrupted: There is no access to some files or the files are occupied by another process." + Environment.NewLine);
                BeginInvoke(new MyDelegate(RefreshPB));
                label1.Text = "Error";
                CheckBtnParsing = false;
                oAuthToolStripMenuItem.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                pathToolStripMenuItem.Enabled = true;
            }
            else
            {
                label1.Text = (string)e.Result;
                CheckBtnParsing = false;
                CheckBtnSave = true;
                oAuthToolStripMenuItem.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                pathToolStripMenuItem.Enabled = true;
            }
        }
        void DocBuild_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }
        void RefreshPB()
        {
            progressBar1.Value = 0;
        }
    }
}
