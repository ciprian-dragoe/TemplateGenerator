using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using TemplateGenerator;
using System.Windows.Forms;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void browseDbfButton_Click(object sender, RoutedEventArgs e)
        {
            // Create an instance of the open file dialog box.
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog.Filter = "DBF Files (.dbf)|*.dbf|All Files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Title = "Selecteaza fisierul DBF";
            openFileDialog.Multiselect = false;
            openFileDialog.ShowDialog();
            if (openFileDialog.FileName != "")
            {
                dbfFolderPath.Text = openFileDialog.FileName;
            }
        }

        private void browseDocxButton_Click(object sender, RoutedEventArgs e)
        {
            // Create an instance of the open file dialog box.
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog.Filter = "DocX Files (.docx)|*.docx|All Files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Title = "Selecteaza fisierul template docx";
            openFileDialog.Multiselect = false;
            openFileDialog.ShowDialog();
            if (openFileDialog.FileName != "")
            {
                docxFolderPath.Text = openFileDialog.FileName;
                if (generatedDocxPath.Text == "")
                {
                    // remove the file.docx from path\file.docx
                    generatedDocxPath.Text = openFileDialog.FileName.Substring(0, openFileDialog.FileName.LastIndexOf("\\"));
                }
            }
            
        }

        private void browseGeneratedDocxButton_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (fbd.SelectedPath != "")
            {
                generatedDocxPath.Text = fbd.SelectedPath;
            }
        }

        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            string[] paths = new string[] { dbfFolderPath.Text, docxFolderPath.Text, generatedDocxPath.Text};
            try
            {
                TemplateGenerator.TemplateManager.Main(paths);
            }
            catch (Exception exceptie)
            {
                System.Windows.MessageBox.Show(exceptie.Message);
            }
        }
    }
}
