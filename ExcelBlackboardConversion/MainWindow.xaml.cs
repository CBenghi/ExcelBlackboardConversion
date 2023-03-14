using System;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using Path = System.IO.Path;

namespace ExcelBlackboardConversion
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

        private void button_Click(object sender, RoutedEventArgs e)
        {
            var di = new DirectoryInfo(txtFolderName.Text); 
            if (!di.Exists) 
            {
                System.Windows.MessageBox.Show("Invalid folder name", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            foreach (var file in di.GetFiles("*.xlsx"))
            {
                var reportName = Path.ChangeExtension(file.FullName, "html");
                if (File.Exists(reportName))
                    File.Delete(reportName);
                var result = Results.FromFile(file);
                if (result != null) 
                {
                    result.CleanUp();
                    Reporting.ReportByQuestion(result, reportName);
                }
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            var startFolder = 
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                + Path.DirectorySeparatorChar;
            if (!string.IsNullOrEmpty(txtFolderName.Text) && System.IO.Directory.Exists(txtFolderName.Text))
            {
                startFolder = txtFolderName.Text;
            }
            using var dialog = new FolderBrowserDialog
            {
                Description = "Select the folder where xlsx files are",
                UseDescriptionForTitle = true,
                SelectedPath = startFolder,
                ShowNewFolderButton = false
            };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtFolderName.Text = dialog.SelectedPath;
            }
        }
    }
}
