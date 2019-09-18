namespace ProjectMigrationHelper
{
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

                var solution = new DteSolution(dte, new OutputWindowTracer(ToolWindowCommand.Instance.VsOutputWindow));
                var folder = solution.Folder;
                if (folder == null)
                    return;

                var fingerPrints = solution.CreateFingerprints();

                foreach (var fingerPrint in fingerPrints)
                {
                    var fileName = Path.Combine(folder, $"{fingerPrint.Key}.{FileNameSuffix.Text}.json");

                    File.WriteAllText(fileName, fingerPrint.Value.ToString());
                }
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }
    }
}