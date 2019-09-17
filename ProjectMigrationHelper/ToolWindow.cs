namespace ProjectMigrationHelper
{
    using System.Runtime.InteropServices;

    using Microsoft.VisualStudio.Shell;

    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    /// </summary>
    [Guid("f3c48ab9-34a6-45ea-9ab1-ab80d87f1e3f")]
    public class ToolWindow : ToolWindowPane
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindow"/> class.
        /// </summary>
        public ToolWindow() : base(null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Caption = "Project Migration Helper";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            Content = new ToolWindowControl();
        }
    }
}