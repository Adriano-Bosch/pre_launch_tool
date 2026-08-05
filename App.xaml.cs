using System.Configuration;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace Pre_Launch_Tool
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private const string AppLogFolderName = "Pre_Launch_Tool\\Logs";

        protected override void OnStartup(StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            DispatcherUnhandledException += App_DispatcherUnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
            base.OnStartup(e);
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            string logPath = WriteFatalLog(e.Exception, "DispatcherUnhandledException");
            ShowFatalError(e.Exception, logPath);
            e.Handled = true;
            Shutdown(-1);
        }

        private void CurrentDomain_UnhandledException(object? sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception ?? new Exception("Unknown unhandled exception");
            string logPath = WriteFatalLog(ex, "CurrentDomainUnhandledException");
            ShowFatalError(ex, logPath);
        }

        private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            string logPath = WriteFatalLog(e.Exception, "TaskSchedulerUnobservedTaskException");
            ShowFatalError(e.Exception, logPath);
            e.SetObserved();
        }

        private static string WriteFatalLog(Exception ex, string source)
        {
            try
            {
                string baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string logFolder = Path.Combine(baseFolder, AppLogFolderName);
                Directory.CreateDirectory(logFolder);

                string logPath = Path.Combine(logFolder, $"fatal_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                var sb = new StringBuilder();
                sb.AppendLine($"Timestamp: {DateTime.Now:O}");
                sb.AppendLine($"Source: {source}");
                sb.AppendLine($"OS: {Environment.OSVersion}");
                sb.AppendLine($"64-bit OS: {Environment.Is64BitOperatingSystem}");
                sb.AppendLine($"64-bit process: {Environment.Is64BitProcess}");
                sb.AppendLine($"Machine: {Environment.MachineName}");
                sb.AppendLine($"User: {Environment.UserName}");
                sb.AppendLine("Exception:");
                sb.AppendLine(ex.ToString());
                File.WriteAllText(logPath, sb.ToString(), Encoding.UTF8);
                return logPath;
            }
            catch
            {
                return "Could not write log file";
            }
        }

        private static void ShowFatalError(Exception ex, string logPath)
        {
            try
            {
                MessageBox.Show(
                    "A critical startup error occurred.\n\n" +
                    $"Error: {ex.Message}\n\n" +
                    "A diagnostic log was saved at:\n" +
                    logPath,
                    "Pre Launch Tool - Startup Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            catch
            {
                // Do not throw from global exception handlers.
            }
        }
    }

}
