using System;
using System.IO;
using System.Windows;
using MahApps.Metro.Controls;
using Microsoft.Win32;
using WPFFolderBrowser;

namespace PTO_Emailer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
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
                if (Properties.Settings.Default.IsFirstRun)
                {
                    SetDefaultFolderPath(Path.GetDirectoryName(xmlFileDialog.FileName));
                    Properties.Settings.Default.IsFirstRun = false;
                }
                CheckFileType(xmlFileDialog.FileName);
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
                
            ReadVacationXML(file);
        }

        private void ReadVacationXML(string file)
        {
            Console.WriteLine(file + " is now being read.");
        }

        private void SetDefaultFolderPath(string path)
        {
            MessageBoxResult dialogResult = MessageBox.Show("Would you like to set " + 
                path + " as the default directory when locating the vacation XML file?", 
                "Set Default XML Directory?", MessageBoxButton.YesNo);
            if (dialogResult == MessageBoxResult.Yes)
            {
                Properties.Settings.Default.InitialPath = path;
            }
            else if (dialogResult == MessageBoxResult.No)
            {
                //do something else
            }
        }

        private void BrowseForFolder(object sender, RoutedEventArgs e)
        {
            WPFFolderBrowserDialog folderBrowser = new WPFFolderBrowserDialog();
            folderBrowser.InitialDirectory = Properties.Settings.Default.InitialPath;
            folderBrowser.Title = "Select default browsing directory";

            var result = folderBrowser.ShowDialog();          
            if (! result.Value)
            {
                SetDefaultFolderPath(result.Value.ToString());
            }
        }

    }
}
