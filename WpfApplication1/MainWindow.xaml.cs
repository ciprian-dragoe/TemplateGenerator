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
using System.ComponentModel;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private BackgroundWorker backgroundGenerateDocx = new BackgroundWorker();

        public MainWindow()
        {
            InitializeComponent();
            dbfFolderPath.Text = @"D:\CLOUD-SYNC\Cipi\QSYNC\6_PROJECTS\TemplateGenerator\WorkingFiles\SATNETTE.DBF";
            docxFolderPath.Text = @"D:\CLOUD-SYNC\Cipi\QSYNC\6_PROJECTS\TemplateGenerator\WorkingFiles\Contract Dolce Sport Mansat -7.docx";
            generatedDocxPath.Text = @"c:\Users\cipri\Desktop\temp\";
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
            backgroundGenerateDocx.WorkerSupportsCancellation = true;
            //backgroundGenerateDocx.WorkerReportsProgress = true;
            backgroundGenerateDocx.DoWork += new DoWorkEventHandler(backgroundGenerateDocx_DoWork);
            backgroundGenerateDocx.RunWorkerAsync(paths);
            cancelButton.IsEnabled = true;
            startButton.IsEnabled = false;
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            backgroundGenerateDocx.CancelAsync();
            cancelButton.IsEnabled = false;
            startButton.IsEnabled = true;
            
        }

        private void backgroundGenerateDocx_DoWork(object o, DoWorkEventArgs e)
        {
            BackgroundWorker bw = o as BackgroundWorker;
            try
            {
                foreach (var generatedFileName in TemplateManager.generateDocxTemplates(((string[])e.Argument)[0], ((string[])e.Argument)[1], ((string[])e.Argument)[2]))
                {
                    if (bw.CancellationPending == true)
                    {
                        e.Cancel = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
                e.Cancel = true;
                this.Dispatcher.Invoke((Action)(() => cancelButton_Click(new object(), new RoutedEventArgs())));
            }
        }   // backgroundGenerateDocx_DoWork
    }
}
