﻿using System;
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

namespace PTO_Emailer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        ArrayList employees = new ArrayList();

        public MainWindow()
        {
            InitializeComponent();
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
                //pass
            }
            else if (file.Substring(file.Length - 3).ToUpper() == "XLS")
            {
                File.Copy(file, file.Substring(0, file.Length - 3) + "xml");
                file = file.Substring(0, file.Length - 3) + "xml";
            }
            else
            {
                MessageBox.Show("Not a valid filetype. Please try again.");
                return;
            }
            //clear any old data
            employees.Clear();
            EmployeeComboBox.Items.Clear();

            ReadVacationXML(file);
            BindEmployeeDataToComboBox();
            EnableControls();
        }


        private void EnableControls()
        {
            CreateMailingsTab.Visibility = Visibility.Visible;
            MailButton.Visibility = Visibility.Visible;
            EmployeeComboBox.Visibility = Visibility.Visible;
        }


        private void ReadVacationXML(string file)
        {
            Console.WriteLine(file + " is now being read.");

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(file);
            XmlNodeList Rows = xmlDoc.GetElementsByTagName("Row");
            string balanceColumn = "";
            int i = 1;
            EmployeeData emp = new EmployeeData();

            foreach (XmlNode row in Rows)
            {
                if (XmlParser.IsAttributeName(row.Attributes, "ss:Height", "10.99"))
                {
                    if (emp.Name.Equals(""))
                    {
                        emp.Name = XmlParser.FindRowColData(row, "4");
                    }
                }

                if (!emp.Name.Equals(""))
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
                        if (!emp.Name.Equals("") && !emp.Vacation.Equals("") && !emp.Sick.Equals(""))
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
                i++;
            }
        }


        private string ConvertData(string charStr)
        {
            string convertedStr = "";
            string tempStr = "";
            string[] strArray = charStr.Split('#');
            bool onesDigitIndex = false;

            int i = 1;
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
                catch
                {
                    //Console.WriteLine(e.ToString());
                }
                i++;
            }
            return convertedStr;
        }


        private void BindEmployeeDataToComboBox()
        {
            ComboBoxItem myFirstItem = new ComboBoxItem
            {
                Content = "All Employees",
                IsSelected = true
            };
            EmployeeComboBox.Items.Add(myFirstItem);

            foreach (EmployeeData employee in employees)
            {
                EmployeeComboBox.Items.Add(employee.Name);
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
            StatusLabel.Text = "";
        }


        private void MailButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            StatusLabel.Text = MailButton.Tag.ToString();
        }


        private void MailButton_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            StatusLabel.Text = "";
        }


        private void MailButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine(EmployeeComboBox.SelectedItem.ToString());
            if (EmployeeComboBox.SelectedItem.ToString().Equals("System.Windows.Controls.ComboBoxItem: All Employees"))
            {
                foreach (EmployeeData employee in employees)
                {
                    CreateEmail(employee.Name, employee.Vacation, employee.Sick);
                }
            }
            else
            {
                foreach (EmployeeData employee in employees)
                {
                    if (EmployeeComboBox.SelectedItem.ToString().Equals(employee.Name))
                    {
                        CreateEmail(employee.Name, employee.Vacation, employee.Sick);
                        break;
                    }
                }
            }
            
        }


        private void CreateEmail(string recipient, string vacation, string sick)
        {
            string[] empName = recipient.Split(',');
            
            string bodyStr = "Dear " + empName[1] + "," + "\r\n" +
                "\r\n" + "Your current vacation balance is " + vacation + " hours." +
                "\r\n" + "Your current sick balance is " + sick + " hours.";
            string TO_Recipients = recipient;
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
