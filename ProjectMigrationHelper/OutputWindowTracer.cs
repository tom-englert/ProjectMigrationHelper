namespace ProjectMigrationHelper
{
    using System;

    using JetBrains.Annotations;

    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Shell.Interop;

    public class OutputWindowTracer : ITracer
    {
        private readonly IVsOutputWindow _outputWindow;

        private static Guid _outputPaneGuid = new Guid("{51DB7742-B447-4BF8-B62F-D82FE1B3B848}");

        public OutputWindowTracer(IVsOutputWindow outputWindow)
        {
            _outputWindow = outputWindow;
        }

        private void LogMessageToOutputWindow([CanBeNull] string value)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                var errorCode = _outputWindow.GetPane(ref _outputPaneGuid, out var pane);

                if (ErrorHandler.Failed(errorCode) || pane == null)
                {
                    _outputWindow.CreatePane(ref _outputPaneGuid, "Project Migration Helper", Convert.ToInt32(true), Convert.ToInt32(false));
                    _outputWindow.GetPane(ref _outputPaneGuid, out pane);
                }

                pane?.OutputString(value);
            });
        }

        public void TraceError(string value)
        {
            WriteLine(string.Concat("Error", @" ", value));
        }

        public void TraceWarning(string value)
        {
            WriteLine(string.Concat("Warning", @" ", value));
        }

        public void WriteLine(string value)
        {
            LogMessageToOutputWindow(value + Environment.NewLine);
        }
    }
}
