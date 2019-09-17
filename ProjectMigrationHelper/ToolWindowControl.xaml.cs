namespace ProjectMigrationHelper
{
    using System;
    using System.IO;
    using System.Windows;
    using System.Windows.Input;

    using Microsoft.Win32;

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
            Cursor = Cursors.Wait;

            try
            {
                var dte = ToolWindowCommand.Instance.Dte;

                var solution = new DteSolution(dte, new OutputWindowTracer(ToolWindowCommand.Instance.VsOutputWindow));

                var fingerPrint = solution.CreateFingerprint().ToString();

                var dlg = new SaveFileDialog
                {
                    AddExtension = true,
                    CheckPathExists = true,
                    DefaultExt = ".json",
                    FileName = DateTime.Now.ToString("yy-MM-dd hh-mm-ss"),
                    InitialDirectory = solution.Folder
                };

                if (dlg.ShowDialog() == true)
                {
                    File.WriteAllText(dlg.FileName, fingerPrint);
                }
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }
    }
}