using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading;

namespace PI_UFL_ini_bat_file_creator.Wpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        // initializing streamwriter for log file
        static string logFile = "logFile" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
        StreamWriter loggingWriter = new StreamWriter(logFile, true);
        // initializing the arrays for reading the excel data
        string[] column001Data; // position in line column
        string[] column002Data; // needs to be logged column
        string[] column003Data; // tagname column
        string[] column004Data; // configuration data type column
        string[] column005Data; // description data column


        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            cmbDataType.Items.Add("ASCII");
            cmbDataType.Items.Add("Serial");
            cmbDataSource.Items.Add("PLC");
            cmbDataSource.Items.Add("DTPS 1");
            cmbDataSource.Items.Add("DTPS 2");
            cmbDataSource.Items.Add("ENGINE ROOM");
            cmbDataSource.Items.Add("ENGINE ROOM AMS");
            cmbDataSource.Items.Add("ENGINE ROOM FUEL");
            cmbDataType.Text = "ASCII";
            cmbDataSource.Text = "PLC";
        }

        private void btnExcelFile_Click(object sender, RoutedEventArgs e)
        {
            lblStatusBar.Content = "Reading the excel file...";
            Microsoft.Win32.OpenFileDialog excelOpenFileDialog = new Microsoft.Win32.OpenFileDialog();
            excelOpenFileDialog.DefaultExt = ".xlsx";
            excelOpenFileDialog.Multiselect = false;
            excelOpenFileDialog.Filter = "Excel files (.xlsx)|*.xlsx";

            bool? excelOpenFileDialogOpen = excelOpenFileDialog.ShowDialog();
            if (excelOpenFileDialogOpen == true)
            {
                txbExcelFile.Text = excelOpenFileDialog.FileName;
                try
                {
                    string[][] excelData = ReadDataFromExcelFile(excelOpenFileDialog.FileName);
                    column001Data = excelData[0];
                    column002Data = excelData[1];
                    column003Data = excelData[2];
                    column004Data = excelData[3];
                    column005Data = excelData[4];
                    lblStatusBar.Content = "Done reading the excel file";
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Something went wrong during the reading of the excel file");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Check the ReadDataFromExcelFile method");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    lblStatusBar.Content = "Something went wrong reading the excel file";
                    System.Windows.MessageBox.Show("Something went wrong during the reading of the excel file.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }


        private void btnOutput_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog pathDialog = new FolderBrowserDialog();
            DialogResult pathDialogOpen = pathDialog.ShowDialog();
            if (pathDialogOpen == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(pathDialog.SelectedPath))
            {
                txtOutput.Text = pathDialog.SelectedPath.ToString();
            }
        }

        private void btnCreateFiles_Click(object sender, RoutedEventArgs e)
        {

            #region register user input
            string unitCode = txtUnitCode.Text;
            string unitName = txtUnitName.Text;
            string dataType = cmbDataType.Text;
            string dataSource = cmbDataSource.Text;
            string filePathExcel = txbExcelFile.Text;
            string filePathOutput = txtOutput.Text;
            #endregion register user input

            #region check if input is ok
            // unit code and unit name is filled in
            if (String.IsNullOrEmpty(unitCode) || String.IsNullOrEmpty(unitName))
            {
                System.Windows.MessageBox.Show("Please file in the unit code and/or unit name.\nThe ini file and bat file will not be created properly.",
                    "Excel input", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            // check if data type and data source are chosen
            if (String.IsNullOrEmpty(dataType) || String.IsNullOrEmpty(dataSource))
            {
                System.Windows.MessageBox.Show("Please choose the data type and data source.\nThe ini file and bat file will not be created properly.",
                    "Excel input", MessageBoxButton.OK, MessageBoxImage.Warning);
            }


            // excel file is chosen
            bool filePathExcelOk = false;
            if (String.IsNullOrEmpty(filePathExcel))
            {
                System.Windows.MessageBox.Show("Please choose an excel file as input.\nThe ini file and bat file can not be created.",
                    "Excel input", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                filePathExcelOk = true;
            }

            // output file is chosen
            bool filePathOutputOk = false;
            if (String.IsNullOrEmpty(filePathOutput))
            {
                System.Windows.MessageBox.Show("Please choose an output file as input.\nThe ini file and bat file can not be created.",
                    "File output", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                filePathOutputOk = true;
            }
            #endregion check if input is ok

            // if all checks are ok we can do something
            if (filePathExcelOk && filePathOutputOk)
            {
                // start writing the ini file and the bat file
                // initializing streamwriter for ini file
                string dataPrefix;
                switch (dataSource)
                {
                    case "PLC":
                        dataPrefix = "PLC";
                        break;
                    case "DTPS 1":
                        dataPrefix = "DTPS1";
                        break;
                    case "DTPS 2":
                        dataPrefix = "DTPS2";
                        break;
                    case "ENGINE ROOM":
                        dataPrefix = "ER";
                        break;
                    case "ENGINE ROOM AMS":
                        dataPrefix = "AMS";
                        break;
                    case "ENGINE ROOM FUEL":
                        dataPrefix = "FUEL";
                        break;
                    default:
                        dataPrefix = "";
                        break;
                }

                #region writing the ini effective file
                StreamWriter iniFileWriter;
                try
                {
                    iniFileWriter = new StreamWriter(filePathOutput + "\\" + unitCode + "_" + "UFL" + "_" + dataPrefix + ".ini", true);
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " The file for the output of the ini file was too long. It was written in the MyDocuments folder.");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    iniFileWriter = new StreamWriter(Environment.SpecialFolder.MyDocuments + "\\" + unitCode + "_" + "UFL" + "_" + dataPrefix + ".ini", true);
                }


                // writing the header of the ini file
                try
                {
                    WritingTheHeader(iniFileWriter, unitCode, unitName, dataSource, dataType, dataPrefix);
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Something went wrong during the writing of the header of the ini file");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Check the WritingTheHeader method");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    System.Windows.MessageBox.Show("Something went wrong during the writing of the header of the ini file. Check the log file for more information."
                        , "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                // writing the log part of the ini file
                try
                {
                    WritingTheLoggingPart(iniFileWriter, unitCode, dataPrefix);
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Something went wrong during the writing of the log part of the ini file");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Check the WritingTheLoggingPart method");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    System.Windows.MessageBox.Show("Something went wrong during the writing of the log part of the ini file. Check the log file for more information."
                        , "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                int amountOfTags = column001Data.Length - 2;//int.Parse(column001Data[column001Data.Length - 1]);

                // writing the field part of the ini file
                try
                {
                    WritingTheFieldPart(iniFileWriter, amountOfTags);
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Something went wrong during the writing of the field part of the ini file");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Check the WritingTheFieldPart method");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    System.Windows.MessageBox.Show("Something went wrong during the writing of the field part of the ini file. Check the log file for more information."
                        , "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                // writing the MSG first line part
                try
                {
                    WritingTheMSGFirstLinePart(iniFileWriter, dataPrefix);
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Something went wrong during the writing of the msg first line part of the ini file");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Check the WritingTheMSGFirstLinePart method");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    System.Windows.MessageBox.Show("Something went wrong during the writing of the msg first line part of the ini file. Check the log file for more information."
                        , "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                // writing the MSG part (the starts)
                try
                {
                    WritingTheMSGPart(iniFileWriter, dataPrefix);
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Something went wrong during the writing of the msg part of the ini file");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Check the WritingTheMSGPart method");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    System.Windows.MessageBox.Show("Something went wrong during the writing of the msg part of the ini file. Check the log file for more information."
                        , "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                // writing the build tag names part
                try
                {
                    WritingTheBuildTagNamesPart(iniFileWriter, amountOfTags);
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Something went wrong during the writing of the building tags part of the ini file");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Check the WritingTheBuildTagNamesPart method");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    System.Windows.MessageBox.Show("Something went wrong during the writing of the building tags part of the ini file. Check the log file for more information."
                        , "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                //
                try
                {
                    WritingDataToPIPart(iniFileWriter, amountOfTags);
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Something went wrong during the writing of the write data to PI part of the ini file");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Check the WritingDataToPIPart method");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    System.Windows.MessageBox.Show("Something went wrong during the writing of the write data to PI part of the ini file. Check the log file for more information."
                        , "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                iniFileWriter.Close();
                iniFileWriter.Dispose();
                #endregion writing the effective file

                #region writing the bat effective file
                StreamWriter batFileWriter;
                try
                {
                    batFileWriter = new StreamWriter(filePathOutput + "\\" + unitCode + "_" + "UFL" + "_" + dataPrefix + ".bat", true);
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " The file for the output of the bat file was too long. It was written in the MyDocuments folder.");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    batFileWriter = new StreamWriter(Environment.SpecialFolder.MyDocuments + "\\" + unitCode + "_" + "UFL" + "_" + dataPrefix + ".bat", true);
                    System.Windows.MessageBox.Show(" The file for the output of the bat file was too long. It was written in the MyDocuments folder."
                        , "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                try
                {
                    WritingBatFile(batFileWriter, unitCode, dataPrefix, dataSource);
                }
                catch (Exception ex)
                {
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Something went wrong during the writing of the bat file");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + " Check the WritingBatFile method");
                    loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
                    System.Windows.MessageBox.Show("Something went wrong during the writing of the bat file", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                batFileWriter.Close();
                batFileWriter.Dispose();
                #endregion writing the bat effective file

                System.Windows.MessageBox.Show("Finish creating the ini and bat file.\nIf something is wrong in this file then check the log file.", "ini bat file creation", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                loggingWriter.WriteLine(DateTime.Now.ToString() + " The path checks failed. No ini or bat file was created");
                System.Windows.MessageBox.Show("The path checks failed. No ini or bat file was created", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // if the log file is empty it will be deleted automatically when closing the program
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if ((new FileInfo(logFile).Length == 0) && (new FileInfo(logFile).Exists))
            {
                loggingWriter.Close();
                loggingWriter.Dispose();
                File.Delete(logFile);
            }
        }

        #region definition of the methods
        void WritingTheHeader(StreamWriter iniFileWriter, string unitCode, string unitName, string dataSource, string dataType, string dataPrefix)
        {
            string serviceName = WriteServiceName(unitCode, dataSource);
            iniFileWriter.WriteLine("'UFL service name : " + serviceName);
            iniFileWriter.WriteLine($"'Unit : {unitCode} {unitName}");
            iniFileWriter.WriteLine($"'Device : {dataSource}");
            iniFileWriter.Write(System.Environment.NewLine);
            iniFileWriter.WriteLine("[INTERFACE]");
            if (dataType == "ASCII")
            {
                iniFileWriter.WriteLine("PLUG-IN = AsciiFiles.dll");
                iniFileWriter.Write(System.Environment.NewLine);
                iniFileWriter.WriteLine("[PLUG-IN]");
                iniFileWriter.WriteLine($"IFM = D:\\interfaceData\\{unitCode}\\{dataPrefix}\\{dataPrefix}_{unitCode}*.txt");
                iniFileWriter.WriteLine("ERR = _BAD");
                iniFileWriter.WriteLine("IFS = C");           // chronologuous reading the txt files
                iniFileWriter.WriteLine("NEWLINE = 13,10");   // ASCII code for <CR> carriage return <LF> line feed
                iniFileWriter.WriteLine("PURGETIME = 5s");
                iniFileWriter.WriteLine("REN = _OK");
                iniFileWriter.WriteLine("PFN = TRUE");
                iniFileWriter.WriteLine("PFN_PREFIX = DATAFILE_*");
            }
            else if (dataType == "Serial")
            {
                iniFileWriter.WriteLine("PLUG-IN=Serial.dll");
                iniFileWriter.Write(System.Environment.NewLine);
                iniFileWriter.WriteLine("[PLUG-IN]");
                iniFileWriter.WriteLine("BITS = 8");
                iniFileWriter.WriteLine("COM = %PORT%");
                iniFileWriter.WriteLine("PARITY = NO");
                iniFileWriter.WriteLine("SPEED = 9600");
                iniFileWriter.WriteLine("STOPBITS = 0");
                iniFileWriter.WriteLine("NEWLINE = 13,10");   // ASCII code for <CR> carriage return <LF> line feed
            }
        }

        void WritingTheLoggingPart(StreamWriter iniFileWriter, string unitCode, string dataPrefix)
        {
            iniFileWriter.Write(System.Environment.NewLine);
            iniFileWriter.WriteLine("[SETTING]");
            iniFileWriter.WriteLine("' Debug Levels");
            iniFileWriter.WriteLine("'0 (Default) No debug output.");
            iniFileWriter.WriteLine("'1 Tasks that are normally performed once, such as startup and shutdown messages, points added to the interface's cache, etc.");
            iniFileWriter.WriteLine("'2 More detail than level 1.");
            iniFileWriter.WriteLine("'3 Include raw data.");
            iniFileWriter.WriteLine("'4 Include data about to be sent to PI server.");
            iniFileWriter.WriteLine("'5 Include read scan cycles start and end time, interface internal cache refresh cycles starts and ends times, etc.");
            iniFileWriter.WriteLine("'6 Log lines read from input before processing.");
            iniFileWriter.WriteLine("DEB = 1"); // debug level setting
            iniFileWriter.WriteLine($"MSGINERROR = D:\\Logs\\interfaces\\{unitCode}_{dataPrefix}_error.txt");
            iniFileWriter.WriteLine($"OUTPUT = D:\\Logs\\interfaces\\{unitCode}_{dataPrefix}_log.txt");
            iniFileWriter.WriteLine("MAXLOG = 10");
            iniFileWriter.WriteLine("MAXLOGSIZE = 100");
            iniFileWriter.WriteLine("LOCALE = en-us");
        }

        string[][] ReadDataFromExcelFile(string excelFileName)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excelFileName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            string[][] dataOutputArray = new string[5][];
            for (int i = 0; i < 5; i++)
            {
                dataOutputArray[i] = WriteDataToArray(rowCount, i + 1, xlRange);
            }

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return dataOutputArray;
        }

        string[] WriteDataToArray(int rowCount, int columnNumber, Excel.Range xlRange)
        {
            string[] output = new string[rowCount - 1];
            int counter = 0;
            // remove the header so start counting from 2 in the excel (excel starts counting from 1)
            for (int i = 2; i <= rowCount; i++)
            {
                // start writing everything in column 1
                if (xlRange.Cells[i, columnNumber] != null && xlRange.Cells[i, columnNumber].Value != null)
                {
                    output[counter] = xlRange.Cells[i, columnNumber].Value.ToString();
                    counter++;
                }
            }
            return output;
        }

        void WritingTheFieldPart(StreamWriter iniFileWriter, int amountOfTags)
        {
            int fieldCounter = 0;
            int stringCounter = 2;
            int floatCounter = 0;
            iniFileWriter.Write(System.Environment.NewLine);
            iniFileWriter.WriteLine("[FIELD]");

            for (int i = 0; i < amountOfTags; i++)
            {
                iniFileWriter.WriteLine($"FIELD({i + 1}).Name = \"Tag{i + 1}\"");
                fieldCounter++;
                if (column004Data[i] == "String")
                {
                    stringCounter++;
                }
                else if (column004Data[i] == "Float")
                {
                    floatCounter++;
                }
                else if (column004Data[i] == "Substring")
                {
                    stringCounter += 2;
                }
            }

            for (int i = 1; i <= floatCounter; i++)
            {
                fieldCounter++;
                iniFileWriter.WriteLine($"FIELD({fieldCounter}).Name = \"Flo{i}\"");
                iniFileWriter.WriteLine($"FIELD({fieldCounter}).Type = \"Number\"");
            }

            for (int i = 1; i <= stringCounter; i++)
            {
                fieldCounter++;
                iniFileWriter.WriteLine($"FIELD({fieldCounter}).Name = \"Str{i}\"");
            }

            fieldCounter++;
            iniFileWriter.WriteLine($"FIELD({fieldCounter}).Name = \"TimeStampTmp1\"");
            iniFileWriter.WriteLine($"FIELD({fieldCounter}).Type = \"DateTime\"");
            iniFileWriter.WriteLine($"FIELD({fieldCounter}).Format = \"yyyy/MM/dd hh:mm:ss.nnn\"");
            fieldCounter++;
            iniFileWriter.WriteLine($"FIELD({fieldCounter}).Name = \"UnitName\"");
        }

        void WritingTheMSGFirstLinePart(StreamWriter iniFileWriter, string dataPrefix)
        {
            iniFileWriter.Write(System.Environment.NewLine);
            iniFileWriter.WriteLine("[MSG]");
            iniFileWriter.WriteLine("MSG(1).Name = \"MSG_FirstLine\"");
            iniFileWriter.WriteLine($"MSG(2).Name = \"MSG_{dataPrefix}\"");
            iniFileWriter.Write(System.Environment.NewLine);
            string msgFirstLine = "[MSG_FirstLine]";
            string msgFirstLineUnder = "'";
            for (int i = 0; i < msgFirstLine.Length - 1; i++)
            {
                msgFirstLineUnder += "-";
            }
            iniFileWriter.WriteLine(msgFirstLine);
            iniFileWriter.WriteLine(msgFirstLineUnder);
            iniFileWriter.WriteLine("MSG_FirstLine.FILTER=C1==\"DATAFILE_ * \"");
            iniFileWriter.Write(System.Environment.NewLine);
            string extraDataFromTheLineSentence = "' Extract data from the line Sentence";
            string extraDataFromTheLineSentenceUnder = "'";
            for (int i = 0; i < extraDataFromTheLineSentence.Length - 1; i++)
            {
                extraDataFromTheLineSentenceUnder += "-";
            }
            iniFileWriter.WriteLine(extraDataFromTheLineSentence);
            iniFileWriter.WriteLine(extraDataFromTheLineSentenceUnder);
            iniFileWriter.WriteLine("UnitName = [\"*_*_(*)_*_*\"]");
        }

        void WritingTheMSGPart(StreamWriter iniFileWriter, string dataPrefix)
        {
            iniFileWriter.Write(System.Environment.NewLine);
            string msgHeader = "[MSG_" + dataPrefix + "]";
            string msgHeaderUnder = "'";
            for (int i = 0; i < msgHeader.Length - 1; i++)
            {
                msgHeaderUnder += "-";
            }
            iniFileWriter.WriteLine(msgHeader);
            iniFileWriter.WriteLine(msgHeaderUnder);
            iniFileWriter.WriteLine($"MSG_PLC.FILTER=C2==\"*{dataPrefix}*\"");
            string starLine = "*";
            string starLineComment = ",*";
            int amountOfStars = 0;
            try
            {
                amountOfStars = int.Parse(column001Data[column001Data.Length - 1]);
            }
            catch (FormatException ex)
            {
                loggingWriter.WriteLine(DateTime.Now.ToString() + " Something went wrong during the writing of the msg part of the ini file");
                loggingWriter.WriteLine(DateTime.Now.ToString() + " The parse of the amountOfStars went wrong in the WritingTheMSGPart method");
                loggingWriter.WriteLine(DateTime.Now.ToString() + ex.StackTrace);
            }

            for (int i = 1; i < amountOfStars; i++)
            {
                starLine += ",*";
            }
            for (int i = 1; i < amountOfStars - 1; i++)
            {
                starLineComment += ",*";
            }
            iniFileWriter.WriteLine($"'MSG_PLC.FILTER=C2==\"????????????????????????*{dataPrefix}*\"" + starLineComment);
            iniFileWriter.Write(System.Environment.NewLine);
            string extractDataFromTheMessageSentence = "'Extract data from the MESSAGE Sentence";
            string extractDataFromTheMessageSentenceUnder = "'";
            for (int i = 0; i < extractDataFromTheMessageSentence.Length - 1; i++)
            {
                extractDataFromTheMessageSentenceUnder += "-";
            }
            iniFileWriter.WriteLine(extractDataFromTheMessageSentence);
            iniFileWriter.WriteLine(extractDataFromTheMessageSentenceUnder);
            iniFileWriter.WriteLine("Str1 = [\"(*)" + starLineComment + "\"] ' Timestamp");
            iniFileWriter.WriteLine("Str2 = SUBSTR(Str1, 1, 23)");
            iniFileWriter.WriteLine("TimeStampTmp1 = Str2");

            int stringCounter = 2;
            int subStringCounter = 1;
            int floatCounter = 0;
            for (int i = 0; i < column001Data.Length - 2; i++)
            {
                string starsBefore = "[\"*";
                int positionInLine = int.Parse(column001Data[i]);
                for (int j = 1; j <= positionInLine; j++)
                {
                    starsBefore += ",*";
                }
                string starPositionInLine = ",(*)";
                string starsAfter = "";
                for (int k = positionInLine + 1; k < amountOfStars - 1; k++)
                {
                    starsAfter += ",*";
                }
                string starPart = starsBefore + starPositionInLine + starsAfter + "\"]";

                string dataTypeMSG = column004Data[i + 2];
                string variablePart = "";

                string descriptionPart = $" ' {column005Data[i + 2]}";

                string writePart;
                switch (dataTypeMSG)
                {
                    case "String":
                        stringCounter++;
                        if (column002Data[i + 2] == "1")
                        {
                            variablePart = String.Format("Str{0} = ", stringCounter);
                        }
                        else
                        {
                            variablePart = String.Format("'Str{0} = ", stringCounter);
                        }
                        writePart = variablePart + starPart + descriptionPart;
                        iniFileWriter.WriteLine(writePart);
                        break;
                    case "Substring":
                        stringCounter++;
                        if (column002Data[i + 2] == "1")
                        {
                            variablePart = String.Format("Str{0} = ", stringCounter);
                        }
                        else
                        {
                            variablePart = String.Format("'Str{0} = ", stringCounter);
                        }
                        writePart = variablePart + starPart + descriptionPart;
                        iniFileWriter.WriteLine(writePart);

                        if ((int.Parse(column001Data[i + 1]) == int.Parse(column001Data[i + 2])) && (dataTypeMSG == "Substring"))
                        {
                            subStringCounter++;
                        }
                        else if (int.Parse(column001Data[i + 1]) < int.Parse(column001Data[i + 2]) && (dataTypeMSG == "Substring"))
                        {
                            subStringCounter = 1;
                        }
                        stringCounter++;
                        if (column002Data[i + 2] == "1")
                        {
                            variablePart = String.Format("Str{0} = SUBSTR(Str{1},{2},1)", stringCounter, stringCounter - 1, subStringCounter);
                        }
                        else
                        {
                            variablePart = String.Format("'Str{0} = SUBSTR(Str{1},{2},1)", stringCounter, stringCounter - 1, subStringCounter);
                        }
                        iniFileWriter.WriteLine(variablePart);
                        break;
                    case "Float":
                        floatCounter++;
                        variablePart = String.Format("Flo{0} = ", floatCounter);
                        writePart = variablePart + starPart + descriptionPart;
                        iniFileWriter.WriteLine(writePart);
                        break;
                    default:
                        iniFileWriter.WriteLine("");
                        break;
                }
            }
        }

        void WritingTheBuildTagNamesPart(StreamWriter iniFileWriter, int amountOfTags)
        {
            iniFileWriter.Write(System.Environment.NewLine);
            string calculationPart = "'Calculations";
            string calculationPartUnder = "";
            for (int i = 0; i < calculationPart.Length; i++)
            {
                calculationPartUnder += "-";
            }
            iniFileWriter.WriteLine(calculationPart);
            iniFileWriter.WriteLine(calculationPartUnder);
            iniFileWriter.Write(System.Environment.NewLine);

            string buildTagNamesPart = "'Build tag names";
            string buildTagNamesPartUnder = "";
            for (int i = 0; i < buildTagNamesPart.Length; i++)
            {
                buildTagNamesPartUnder += "-";
            }
            iniFileWriter.WriteLine(buildTagNamesPart);
            iniFileWriter.WriteLine(buildTagNamesPartUnder);
            iniFileWriter.Write(System.Environment.NewLine);

            int floatCounter = 0;
            int stringCounter = 2;
            string variablePart = "";
            for (int i = 1; i <= amountOfTags; i++)
            {
                switch (column004Data[i + 1])
                {
                    case "String":
                        stringCounter++;
                        variablePart = $"' (variable=Str{stringCounter})";
                        break;
                    case "Substring":
                        // add 2 because of str1 = [***] and str2 = substr(str1, part
                        stringCounter += 2;
                        variablePart = $"' (variable=Str{stringCounter})";
                        break;
                    case "Float":
                        floatCounter++;
                        variablePart = $"' (variable=Flo{floatCounter})";
                        break;
                    default:
                        variablePart = "'WRONG INPUT IN THIS CELL IN EXCEL";
                        break;
                }

                if (column002Data[i + 1] == "1")
                {
                    iniFileWriter.WriteLine($"Tag{i} = \"{column003Data[i + 1]} {variablePart}");
                }
                else
                {
                    iniFileWriter.WriteLine($"'Tag{i} = \"{column003Data[i + 1]} {variablePart}");
                }
            }
        }

        void WritingDataToPIPart(StreamWriter iniFileWriter, int amountOfTags)
        {
            iniFileWriter.Write(System.Environment.NewLine);
            string writesDataToPIPart = "'Writes data to PI";
            string writesDataToPIPartUnder = "";
            for (int i = 0; i < writesDataToPIPart.Length; i++)
            {
                writesDataToPIPartUnder += "-";
            }
            iniFileWriter.WriteLine(writesDataToPIPart);
            iniFileWriter.WriteLine(writesDataToPIPartUnder);
            iniFileWriter.Write(System.Environment.NewLine);

            int stringCounter = 2;
            int floatCounter = 0;
            for (int i = 1; i <= amountOfTags; i++)
            {
                string writeTagPart = "";
                switch (column004Data[i + 1])
                {
                    case "String":
                        stringCounter++;
                        if (column002Data[i + 1] == "1")
                        {
                            writeTagPart = $"IF(Str{stringCounter} is NOT NULL) THEN StoreInPI(Tag{i},,TimeStampTmp1, Str{stringCounter},,) ENDIF";
                        }
                        else
                        {
                            writeTagPart = $"'IF(Str{stringCounter} is NOT NULL) THEN StoreInPI(Tag{i},,TimeStampTmp1, Str{stringCounter},,) ENDIF";
                        }
                        break;
                    case "Substring":
                        // add 2 because of str1 = [***] and str2 = substr(str1, part
                        stringCounter++;
                        if (column002Data[i + 1] == "1")
                        {
                            writeTagPart = $"IF(Str{stringCounter} is NOT NULL) THEN StoreInPI(Tag{i},,TimeStampTmp1, Str{stringCounter},,) ENDIF";
                        }
                        else
                        {
                            writeTagPart = $"'IF(Str{stringCounter} is NOT NULL) THEN StoreInPI(Tag{i},, TimeStampTmp1, Str{stringCounter},,) ENDIF";
                        }
                        stringCounter++;
                        break;
                    case "Float":
                        floatCounter++;
                        if (column002Data[i + 1] == "1")
                        {
                            writeTagPart = $"IF(Flo{floatCounter} is NOT NULL) THEN StoreInPI(Tag{i},, TimeStampTmp1, Flo{floatCounter},,) ENDIF";
                        }
                        else
                        {
                            writeTagPart = $"'IF(Flo{floatCounter} is NOT NULL) THEN StoreInPI(Tag{i},, TimeStampTmp1, Flo{floatCounter},,) ENDIF";
                        }
                        break;
                    default:
                        writeTagPart = "";
                        break;
                }
                iniFileWriter.WriteLine(writeTagPart);
            }
        }

        void WritingBatFile(StreamWriter batFileWriter, string unitCode, string dataPrefix, string dataSource)
        {
            batFileWriter.WriteLine($"TITLE {unitCode}_{dataPrefix}");
            string serviceName = WriteServiceName(unitCode, dataSource);
            batFileWriter.WriteLine("\"C:\\Program Files(x86)\\PIPC\\Interfaces\\PI_UFL\\PI_UFL.exe\" " +
                serviceName +
                $" /CF = \"C:\\Program Files (x86)\\PIPC\\Interfaces\\PI_UFL\\INI\\V2\\{unitCode}_UFL_{dataPrefix}.ini\"" +
                $" /UTC /des = 290 /PS = {unitCode}_UFL_{dataPrefix} /host = %DNS_SERVER_NAME%:5450 /f = 00:00:05");
        }

        string WriteServiceName(string unitCode, string dataSource)
        {
            string serviceNameUnitCode = unitCode.Substring(1);
            string serviceNameDataCode;
            switch (dataSource)
            {
                case "DTPS 1":
                    serviceNameDataCode = "01";
                    break;
                case "DTPS 2":
                    serviceNameDataCode = "01";
                    break;
                case "PLC":
                    serviceNameDataCode = "03";
                    break;
                case "ENGINE ROOM":
                    serviceNameDataCode = "04";
                    break;
                case "ENGINE ROOM AMS":
                    serviceNameDataCode = "04";
                    break;
                case "ENGINE ROOM FUEL":
                    serviceNameDataCode = "04";
                    break;
                default:
                    serviceNameDataCode = "00";
                    break;
            }
            string serviceName = serviceNameUnitCode + serviceNameDataCode;
            return serviceName;
        }
        #endregion definition of the methods
    }
}
