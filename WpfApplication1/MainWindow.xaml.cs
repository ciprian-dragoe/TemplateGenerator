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
using System.ComponentModel;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private BackgroundWorker backgroundGenerateDocx;

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
            OpenFileDialog openFileDialog = new OpenFileDialog();

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
            OpenFileDialog openFileDialog = new OpenFileDialog();

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
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            fbd.ShowDialog();
            if (fbd.SelectedPath != "")
            {
                generatedDocxPath.Text = fbd.SelectedPath;
            }
        }

        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            int totalNumberOfRecords = 0;
            backgroundGenerateDocx = new BackgroundWorker();
            backgroundGenerateDocx.WorkerSupportsCancellation = true;
            backgroundGenerateDocx.WorkerReportsProgress = true;

            try
            {
                DBFreader getMaximumElementsProgressBar = new DBFreader(dbfFolderPath.Text);
                totalNumberOfRecords = getMaximumElementsProgressBar.totalNumberElements;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                cancelButton_Click(new object(), new RoutedEventArgs());
                return;
            }
            string[] paths = new string[] { dbfFolderPath.Text, docxFolderPath.Text, generatedDocxPath.Text, totalNumberOfRecords.ToString()};
            backgroundGenerateDocx.DoWork += new DoWorkEventHandler(BackgroundGenerateDocx_DoWork);
            backgroundGenerateDocx.ProgressChanged += new ProgressChangedEventHandler(BackgroundGenerateDocx_ProgressChanged);
            backgroundGenerateDocx.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundGenerateDocx_RunWorkerCompleted);
            cancelButton.IsEnabled = true;
            startButton.IsEnabled = false;
            backgroundGenerateDocx.RunWorkerAsync(paths);
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            backgroundGenerateDocx.CancelAsync();
            cancelButton.IsEnabled = false;
            startButton.IsEnabled = true;
            progressBarGenerateTemplates.Value = 0;
        }

        private void BackgroundGenerateDocx_DoWork(object o, DoWorkEventArgs e)
        {
            BackgroundWorker bw = o as BackgroundWorker;
            int totalNumberElements = Int32.Parse(((string[])e.Argument)[3]);
            try
            {
                int index = 0;
                foreach (var generatedFileName in TemplateManager.generateDocxTemplates(((string[])e.Argument)[0], ((string[])e.Argument)[1], ((string[])e.Argument)[2]))
                {
                    if (bw.CancellationPending == true)
                    {
                        e.Cancel = true;
                        break;
                    }
                    index++;
                    bw.ReportProgress((int)((double)index / totalNumberElements * 100));
                }
            }            catch (Exception ex)
            {
                e.Cancel = true;
                MessageBox.Show(ex.Message);
            }
        }   // backgroundGenerateDocx_DoWork

        private void BackgroundGenerateDocx_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBarGenerateTemplates.Value = e.ProgressPercentage;
        }

        private void BackgroundGenerateDocx_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.cancelButton_Click(new object(), new RoutedEventArgs());
            if (e.Cancelled == false)
            {
                MessageBox.Show("Fisiere generate cu succes.");
            }
        }
    }
}
