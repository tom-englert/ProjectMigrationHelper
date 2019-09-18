namespace ProjectMigrationHelper
{
    using System;
    using System.IO;
    using System.Windows;
    using System.Windows.Input;

    /// <summary>
    /// Interaction logic for ToolWindow1Control.
    /// </summary>
    public partial class ToolWindowControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindowControl"/> class.
        /// </summary>
        public ToolWindowControl()
        {
            InitializeComponent();
        }

        private void CreateFingerprint_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            Cursor = Cursors.Wait;

            try
            {
                var dte = ToolWindowCommand.Instance.Dte;
                var tracer = new OutputWindowTracer(ToolWindowCommand.Instance.VsOutputWindow);

                var solution = new DteSolution(dte, tracer);
                var rootFolder = solution.Folder;
                if (rootFolder == null)
                    return;

                var targetFolder = Path.Combine(rootFolder, SubFolder.Text);

                Directory.CreateDirectory(targetFolder);

                var fingerPrints = solution.CreateFingerprints();

                foreach (var fingerPrint in fingerPrints)
                {
                    var projectName = fingerPrint.Key;
                    var contents = fingerPrint.Value.ToString();
                    if (contents == "{}")
                        continue;

                    var fileName = Path.Combine(targetFolder, projectName + ".json");

                    tracer.WriteLine("Create fingerprint: " + fileName);

                    File.WriteAllText(fileName, contents);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }
    }
}