namespace ProjectMigrationHelper
{
    using System;
    using System.ComponentModel.Design;

    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;

    using IAsyncServiceProvider = Microsoft.VisualStudio.Shell.IAsyncServiceProvider;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ToolWindowCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("0e05aa0e-8b63-428b-a5a1-7095cec42523");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        public readonly AsyncPackage Package;

        public readonly EnvDTE80.DTE2 Dte;

        public readonly IVsOutputWindow VsOutputWindow;

        private ToolWindowCommand(AsyncPackage package, OleMenuCommandService commandService, EnvDTE80.DTE2 dte, IVsOutputWindow vsOutputWindow)
        {
            Package = package;
            Dte = dte;
            VsOutputWindow = vsOutputWindow;

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ToolWindowCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        public IAsyncServiceProvider ServiceProvider => Package;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ToolWindowCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var dte = await package.GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;
            var commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            var vsOutputWindow = await package.GetServiceAsync(typeof(SVsOutputWindow)) as IVsOutputWindow;
            
            // ReSharper disable AssignNullToNotNullAttribute
            Instance = new ToolWindowCommand(package, commandService, dte, vsOutputWindow);
            // ReSharper restore AssignNullToNotNullAttribute
        }

        /// <summary>
        /// Shows the tool window when the menu item is clicked.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            Package.JoinableTaskFactory.RunAsync(async delegate
            {
                var window = await Package.ShowToolWindowAsync(typeof(ToolWindow), 0, true, Package.DisposalToken);

                if (window?.Frame == null)
                {
                    throw new NotSupportedException("Cannot create tool window");
                }

                await Package.ShowToolWindowAsync(typeof(ToolWindow), 0, true, Package.DisposalToken);
            });
        }
    }
}
