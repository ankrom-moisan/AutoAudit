using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System.Collections.Generic;

namespace AutoAudit
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string[] filePaths;
        private ExternalCommandData _commandData;
        private delegate void ProgressBarDelegate();
        public string secondsRemaining;

        /// <summary>
        /// Initialize the window and instantiate Revit Application Data
        /// </summary>
        /// <param name="commandData">Application Data from Revit</param>
        public MainWindow(ExternalCommandData commandData)
        {
            _commandData = commandData;
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        /// <summary>
        /// Main application logic
        /// Executes on "ok" click
        /// opens selected files, audits them, and saves then closes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            // Guard Condition
            if (filePaths.Length == 0) Close();

            // Setup progress bar animation
            progressBar.IsIndeterminate = false;
            progressBar.Maximum = filePaths.Length;
            progressBar.Minimum = 0;

            // required stateful variables for tracking progress
            int numComplete = 0;
            int numRemaining = filePaths.Length;
            bool backupUndeleted = false;
            IList<string> problemFiles = new List<string>();
            
            // create a timer to show time remaining
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            for (int i = 0; i < filePaths.Length; i++)
            {
                string path = filePaths[i];
                secondsRemaining = TimeLeft(stopWatch, numComplete, numRemaining);

                // Update Progress
                progressBar.Dispatcher.Invoke(new ProgressBarDelegate(UpdateProgress),
                    System.Windows.Threading.DispatcherPriority.Background);

                // Process files
                try
                {
                    // Set the "audit on open" option to true
                    OpenOptions opts = new OpenOptions();
                    if (checkAudit.IsChecked == true) opts.Audit = true;

                    // Convert the string filepath to a ModelPath
                    ModelPath mPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(path);

                    // Open the model (initiating upgrade and audit)
                    Document doc = _commandData.Application.Application.OpenDocumentFile(mPath, opts);

                    // Save and close the model
                    doc.Close(true);

                    // Update counters
                    numComplete++;
                    numRemaining--;

                    // Delete the backup file
                    try
                    {
                        string[] splitPath = path.Split('.');
                        string newPath = "";
                        for (int j = 0; j < splitPath.Length - 1; j++)
                        {
                            newPath += splitPath[j] + ".";
                        }
                        newPath += "0001." + splitPath.Last();
                        File.Delete(newPath);
                    }
                    catch (DirectoryNotFoundException)
                    {
                        backupUndeleted = true;
                        continue;
                    }
                    
                }
                catch (Exception)
                {
                    string fileName = path.Split('\\').Last();
                    problemFiles.Add(fileName);
                    continue;
                }
            }

            stopWatch.Stop();
            TaskDialog.Show("AutoAudit", "All audits and upgrades complete!");

            // Show any files that could not be processed or deleted
            if (backupUndeleted) problemFiles.Add("Some backups could not be deleted");
            if (problemFiles.Count != 0)
            {
                ErrorWindow errWindow = new ErrorWindow(problemFiles);
                errWindow.ShowDialog();
            }

            // Return control to main application
            Close();
        }

        /// <summary>
        /// Calculates the estimated time remaining using the average
        /// time per iteration multiplied by the remaining iterations
        /// </summary>
        /// <param name="sw">A stopwatch object</param>
        /// <param name="numComplete">Completed Iterations</param>
        /// <param name="numRemaining">Remaining Iterations</param>
        /// <returns></returns>
        private string TimeLeft(Stopwatch sw, int numComplete, int numRemaining)
        {
            // Guard Condition
            if (numComplete == 0) return "unknown";

            var time = sw.ElapsedMilliseconds;
            var msPerIteration = time / numComplete;
            var msRemaining = numRemaining * msPerIteration;
            float seconds = msRemaining / 1000;
            if (seconds > 60)
            {
                int min = Convert.ToInt32(Math.Floor(seconds / 60));
                int sec = Convert.ToInt32(seconds % 60);
                return min + " min   " + sec + " sec";
            }
            return seconds.ToString() + " sec";
        }

        /// <summary>
        /// Delegate to update progress bar animation in seperate thread
        /// </summary>
        private void UpdateProgress()
        {
            progressBar.Value += 1;
            lblProcessed.Content = "Processed: " + progressBar.Value;
            lblTimeRemaining.Content = "Time Left: ~" + secondsRemaining;
        }

        /// <summary>
        /// Event Handler to open the File Selection Dialog
        /// Saves selections to a string array "filePaths"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void open_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Revit Families (*.rfa)|*.rfa|Revit Projects (*.rvt)|*.rvt";
            if (openFileDialog.ShowDialog() == true)
            {
                filePaths = openFileDialog.FileNames;
                lblFilesSelected.Content = "Files Selected: " + filePaths.Length;
            }
        }
    }
}
