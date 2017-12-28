using System;
using System.Collections;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using MahApps.Metro.Controls;
using Microsoft.Win32;
using WPFFolderBrowser;

using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;
using System.Text.RegularExpressions;
using System.Text;
using System.ComponentModel;
using System.Threading;
using System.Windows.Data;
using System.Windows.Input;

namespace PTO_Emailer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        ArrayList employees = new ArrayList();
        private System.ComponentModel.BackgroundWorker emailsBackgroundWorker = new BackgroundWorker();
        string applicationMessage = "";
        int StartNumber = 1;


        public MainWindow()
        {
            InitializeComponent();
            InitializeBackgroundWorker();
            InitializeDatagrid();

            this.AllowsTransparency = true;
        }


        public struct EmployeeDataStruct
        {
            public int Num { set; get; }
            public string Name { set; get; }
            public double Vaca { set; get; }
            public double Sick { set; get; }
        }


        private void InitializeBackgroundWorker()
        {
            emailsBackgroundWorker.WorkerReportsProgress = true;
            emailsBackgroundWorker.DoWork += new DoWorkEventHandler(EmailsBackgroundWorker_DoWork);
            emailsBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(EmailsBackgroundWorker_RunWorkerCompleted);
            emailsBackgroundWorker.ProgressChanged += new ProgressChangedEventHandler(EmailsBackgroundWorker_ProgressChanged);
        }


        private void InitializeDatagrid()
        {
            DataGridTextColumn NumCol = new DataGridTextColumn();
            DataGridTextColumn NameCol = new DataGridTextColumn();
            DataGridTextColumn VacaCol = new DataGridTextColumn();
            DataGridTextColumn SickCol = new DataGridTextColumn();

            EmployeeDataGrid.Columns.Add(NumCol);
            EmployeeDataGrid.Columns.Add(NameCol);
            EmployeeDataGrid.Columns.Add(VacaCol);
            EmployeeDataGrid.Columns.Add(SickCol);

            NumCol.Binding = new Binding("Num");
            NameCol.Binding = new Binding("Name");
            VacaCol.Binding = new Binding("Vaca");
            SickCol.Binding = new Binding("Sick");

            NumCol.Header = "Num";
            NameCol.Header = "Employee Name";
            VacaCol.Header = "Vacation";
            SickCol.Header = "Sick";
        }


        private void SelectFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog xmlFileDialog = new OpenFileDialog();
            xmlFileDialog.Filter = "XML files (*.xml)|*.xml|XLS files(*.xls)| *.xls";
            xmlFileDialog.Multiselect = false;
            xmlFileDialog.Title = "Select Vacation XML file";
            xmlFileDialog.InitialDirectory = Properties.Settings.Default.InitialPath;

            if (xmlFileDialog.ShowDialog() == true)
            {
                if (Properties.Settings.Default.IsFirstRun)
                {
                    string fileDirectory = Path.GetDirectoryName(xmlFileDialog.FileName);
                    MessageBoxResult dialogResult = MessageBox.Show("Would you like to set " +
                        fileDirectory + " as the default directory when locating the vacation XML file?",
                        "Set Default XML Directory?", MessageBoxButton.YesNo);
                    if (dialogResult == MessageBoxResult.Yes)
                    {
                        Properties.Settings.Default.InitialPath = fileDirectory;
                    }
                    Properties.Settings.Default.IsFirstRun = false;
                }
                CheckFileType(xmlFileDialog.FileName);
            }
        }


        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            Close();
        }


        private void MetroWindow_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] file = (string[])e.Data.GetData(DataFormats.FileDrop);
                Console.WriteLine(file[0] + " was dropped into the application.");
                CheckFileType(file[0]);
            }
        }


        private void CheckFileType(string file)
        {
            if (file.Substring(file.Length - 3).ToUpper() == "XML")
            {
                //Pass
            }
            else if (file.Substring(file.Length - 3).ToUpper() == "XLS")
            {
                File.Copy(file, file.Substring(0, file.Length - 3) + "xml");
                file = file.Substring(0, file.Length - 3) + "xml";
            }
            else
            {
                StatusLabel.Text = "Invalid file.";
                MessageBox.Show("Not a valid filetype. Please try again.");
                return;
            }
            DropFileHereLabel.Visibility = Visibility.Hidden;

            //clear any old data
            StatusLabel.Text = "";
            employees.Clear();
            EmployeeDataGrid.Items.Clear();
            EmployeeComboBox.Items.Clear();

            CheckFileForErroneousData(file);
            ReadVacationXML(file);
            BindEmployeeDataToControls();
            EnableControls();
        }


        private void CheckFileForErroneousData(string file)
        {
            ArrayList newLines = new ArrayList();

            string[] lines = File.ReadAllLines(file);
            File.Delete(file);
            foreach (string line in lines)
            {
                newLines.Add(Regex.Replace(line, "&", "", RegexOptions.Compiled));
            }

            using (FileStream fs = File.Create(file))
            {
                foreach (string newLine in newLines)
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes(newLine);
                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);
                }
            }
        }


        private void EnableControls()
        {
            CreateMailingsTab.Visibility = Visibility.Visible;
            CreateMailingsTab.IsSelected = true;

            StartNumberTextBox.Visibility = Visibility.Visible;
            MailButton.Visibility = Visibility.Visible;
            EmployeeComboBox.Visibility = Visibility.Visible;
            EmployeesDataTab.Visibility = Visibility.Visible;
        }


        private void ReadVacationXML(string file)
        {
            Console.WriteLine(file + " is now being read.");

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(file);
            XmlNodeList Rows = xmlDoc.GetElementsByTagName("Row");
            string balanceColumn = "";

            EmployeeData emp = new EmployeeData();

            foreach (XmlNode row in Rows)
            {
                if (XmlParser.IsAttributeName(row.Attributes, "ss:Height", "10.99"))
                {
                    if (emp.FullName.Equals(""))
                    {
                        emp.FullName = XmlParser.FindRowColData(row, "4");
                    }
                }

                if (!emp.FullName.Equals(""))
                {
                    if (!balanceColumn.Equals(""))
                    {
                        if (XmlParser.IsRowWithFirstChildText(row, "Vacation"))
                        {
                            emp.Vacation = ConvertData(XmlParser.FindRowColData(row, balanceColumn));
                        }
                        else if (XmlParser.IsRowWithFirstChildText(row, "Sick"))
                        {
                            emp.Sick = ConvertData(XmlParser.FindRowColData(row, balanceColumn));
                        }
                        if (!emp.FullName.Equals("") && !emp.Vacation.Equals("") && !emp.Sick.Equals(""))
                        {
                            employees.Add(emp);
                            Console.WriteLine(emp.ToString() + "\r\n");
                            //reset values
                            emp = new EmployeeData();
                            balanceColumn = "";
                        }
                    }
                    else
                    {
                        balanceColumn = XmlParser.FindColumnContainingText(row, "Balance");
                    }
                }
            }
        }


        private string ConvertData(string charStr)
        {
            string convertedStr = "";
            string tempStr = "";
            string[] strArray = charStr.Split('#');
            bool onesDigitIndex = false;

            foreach (string str in strArray)
            {
                if (str.Contains("."))
                {
                    tempStr = str.Substring(0, str.Length - 1);
                    onesDigitIndex = true;
                }
                else
                {
                    tempStr = str;
                }

                try
                {
                    int x = Int32.Parse(tempStr);
                    char c = Convert.ToChar(x);
                    convertedStr += c.ToString();
                    if (onesDigitIndex)
                    {
                        convertedStr += ".";
                        onesDigitIndex = false;
                    }
                }
                catch //(System.Exception e)
                {
                    //Console.WriteLine(e.ToString());
                }
            }
            return convertedStr;
        }


        private void BindEmployeeDataToControls()
        {
            ComboBoxItem myFirstItem = new ComboBoxItem
            {
                Content = "All Employees",
                IsSelected = true
            };
            EmployeeComboBox.Items.Add(myFirstItem);

            int i = 1;
            foreach (EmployeeData employee in employees)
            {
                EmployeeComboBox.Items.Add(employee.FullName); // add employee to combobox

                EmployeeDataGrid.Items.Add(new EmployeeDataStruct // add employee to datagrid
                {
                    Num = i++,
                    Name = employee.FullName,
                    Vaca = double.Parse(employee.Vacation),
                    Sick = double.Parse(employee.Sick)
                });
            }
        }


        private void BrowseForFolder(object sender, RoutedEventArgs e)
        {
            WPFFolderBrowserDialog folderBrowser = new WPFFolderBrowserDialog();
            folderBrowser.InitialDirectory = Properties.Settings.Default.InitialPath;
            folderBrowser.Title = "Select default browsing directory";

            var result = folderBrowser.ShowDialog();
            if (result.Value != false)
            {
                Properties.Settings.Default.InitialPath = folderBrowser.FileName;
            }
        }


        private void EmployeeComboBox_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            StatusLabel.Text = EmployeeComboBox.Tag.ToString();
        }


        private void EmployeeComboBox_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            StatusLabel.Text = applicationMessage;
        }


        private void MailButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            StatusLabel.Text = MailButton.Tag.ToString();
        }


        private void MailButton_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            StatusLabel.Text = applicationMessage;
        }


        private void StartNumberTextBox_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            StatusLabel.Text = StartNumberTextBox.Tag.ToString();
        }


        private void StartNumberTextBox_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            StatusLabel.Text = applicationMessage;
        }


        private void StartNumberTextBox_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Tab)
            {
                StartNumber = ValidateStartNum(StartNumberTextBox.Text);
                StartNumberTextBox.Text = StartNumber.ToString();
            }
        }


        private void StartNumberTextBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            StartNumber = ValidateStartNum(StartNumberTextBox.Text);
            StartNumberTextBox.Text = StartNumber.ToString();
        }


        private int ValidateStartNum(string text)
        {
            int startInt = 1;
            try
            {
                startInt = int.Parse(text);
            }
            catch
            {
                MessageBox.Show("Please enter a number.");
                return 1;
            }

            if (startInt < 1)
            {
                startInt = 1;
            }
            else if (startInt > employees.Count)
            {
                startInt = employees.Count;
            }
            return startInt;
        }

        private void MailButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine(EmployeeComboBox.SelectedItem.ToString());
            if (EmployeeComboBox.SelectedItem.ToString().Equals("System.Windows.Controls.ComboBoxItem: All Employees"))
            {
                applicationMessage = "Creating Mail Items...";
                EmployeeComboBox.IsEnabled = false;
                MailButton.IsEnabled = false;
                OpenMenuItem.IsEnabled = false;
                DefaultDirectoryMenuItem.IsEnabled = false;

                ProgressBar.Visibility = Visibility.Visible;
                TaskbarItemInfo.ProgressState = System.Windows.Shell.TaskbarItemProgressState.Normal;
                emailsBackgroundWorker.RunWorkerAsync();
            }
            else
            {
                foreach (EmployeeData employee in employees)
                {
                    if (EmployeeComboBox.SelectedItem.ToString().Equals(employee.FullName))
                    {
                        CreateEmail(employee);
                        break;
                    }
                }
            }
            
        }


        private void EmailsBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            int i = 1;
            int LastGroupStartNum = StartNumber;
            foreach (EmployeeData employee in employees.GetRange(StartNumber - 1, employees.Count - (StartNumber - 1)))
            {
                Console.WriteLine("Creating email number " + (StartNumber + i - 1) + " for " + employee.FullName);
                //Thread.Sleep(100); //simulating work for testing purposes
                CreateEmail(employee);
                if (i % 10 == 0)
                {
                    MessageBox.Show("Currently working on emails " + LastGroupStartNum + 
                        " through " + (StartNumber + i - 1) + " out of " + employees.Count + " total.");
                    LastGroupStartNum = StartNumber + i;
                }
                i++;
                emailsBackgroundWorker.ReportProgress((int)((double)i / (double)(employees.Count - StartNumber) * 100));
            }
        }


        private void EmailsBackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value = e.ProgressPercentage;
            TaskbarItemInfo.ProgressValue = (double)(e.ProgressPercentage)/ 100;
        }


        private void EmailsBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            EmployeeComboBox.IsEnabled = true;
            MailButton.IsEnabled = true;
            OpenMenuItem.IsEnabled = true;
            DefaultDirectoryMenuItem.IsEnabled = true;
            applicationMessage = "";

            ProgressBar.Value = 100;
            ProgressBar.Visibility = Visibility.Hidden;
            TaskbarItemInfo.ProgressState = System.Windows.Shell.TaskbarItemProgressState.None;
        }


        private void CreateEmail(EmployeeData employee)
        {
            string bodyStr = "Dear " + employee.FirstName + "," + "\r\n" +
                "\r\n" + "Your current vacation balance is " + employee.Vacation + " hours." +
                "\r\n" + "Your current sick balance is " + employee.Sick + " hours.";
            string TO_Recipients = employee.FullName;
            string CC_Recipients = "";
            string subjectStr = "Your Current Vacation Balance";

            OutlookApp otlApp = new OutlookApp();
            MailItem otlNewMail = otlApp.CreateItem(OlItemType.olMailItem);
            Type WshShell = Type.GetTypeFromProgID("WScript.Shell");

            otlNewMail.Display();
            otlNewMail.Subject = subjectStr;
            otlNewMail.To = TO_Recipients;
            otlNewMail.CC = CC_Recipients;
            var objDoc = otlApp.ActiveInspector().WordEditor;
            var objSel = objDoc.Windows(1).Selection;
            objSel.InsertBefore(bodyStr);

            WshShell = null;
            otlNewMail = null;
            otlApp = null;
        }
    }
}
