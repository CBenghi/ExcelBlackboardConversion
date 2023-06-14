using ExcelBlackboardConversion.MarkPeer;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using Path = System.IO.Path;
using MathNet.Numerics.Statistics;

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
            if (!string.IsNullOrEmpty(txtFolderName.Text) && Directory.Exists(txtFolderName.Text))
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

		private void button2_Click(object sender, RoutedEventArgs e)
		{
            string file;
			file = "C:\\Users\\Claudio\\OneDrive - Northumbria University - Production Azure AD\\2022\\MastersResearch\\Sem2\\Marks\\CoordinationV1.xlsx";
			file = "C:\\Users\\Claudio\\OneDrive - Northumbria University - Production Azure AD\\2022\\MastersResearch\\Sem2\\Marks\\CoordinationV2.xlsx";
			var groups = MarkGroup.GetGroups(file).ToList();
			var collection = new PairingCollection();
			foreach (var group in groups)
			{
				group.PopulateCollection(ref collection);
			}
            Table(collection, (x, y) => x.MarkDeltas.Count.ToString(), "count");
			Table(collection, (x, y) => Statistics.Mean(x.GetDeltas(y)).ToString(), "mean");
			Table(collection, (x, y) => Statistics.StandardDeviation(x.GetDeltas(y)).ToString(), "std");

            Debug.Write("{mrkr}\t{Statistics.Mean(deltas)}\t{Statistics.StandardDeviation(deltas)}\t{deltas.Count}\t");
			Debug.Write("{mrkr}\t{Statistics.Mean(diff)}\t{Statistics.StandardDeviation(diff)}\t{diff.Count}\t");
            Debug.WriteLine("");
			foreach (var mrkr in collection.GetAllMarkers())
            {
                var deltas = collection.GetDeltas(mrkr).Select(x => Math.Abs(x)).ToList();
                Debug.Write($"{mrkr}\t{Statistics.Mean(deltas)}\t{Statistics.StandardDeviation(deltas)}\t{deltas.Count}\t");

				var diff = collection.GetDeltas(mrkr).ToList();
				Debug.Write($"{mrkr}\t{Statistics.Mean(diff)}\t{Statistics.StandardDeviation(diff)}\t{diff.Count}\t");
                Debug.WriteLine("");
			}
		}

		private static void Table(PairingCollection collection, Func<MarkPair,string, string> reportFunction, string header)
		{
            var horMarkers = collection.GetAllMarkers().OrderByDescending(x => x).ToList();
			var verMarkers = collection.GetAllMarkers().ToList();
			StringBuilder sb = new StringBuilder();
			sb.Append($"{header}\t");
			foreach (var refMrkr in horMarkers)
			{
				sb.Append($"{refMrkr}\t");
			}
			sb.AppendLine();

            foreach (var verMrkr in verMarkers)
            {
				sb.Append($"{verMrkr}\t");

                foreach (var horMrkr in horMarkers)
                {
					if (verMrkr == horMrkr)
                    {
						sb.Append($"\t");
					}
					else if (collection.TryGet(verMrkr, horMrkr, out var pair))
					{
                        var ret = reportFunction(pair, verMrkr);
                        if (ret == "NaN")
                            ret = "";
						sb.Append($"{ret}\t");
					}
					else
					{
						sb.Append($"\t");
					}
				}
				sb.AppendLine();
			}
			Debug.WriteLine(sb.ToString());
		}
	}
}
