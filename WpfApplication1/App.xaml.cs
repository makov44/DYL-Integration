using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Data;
using DYL.EmailIntegration.Helpers;
using DYL.EmailIntegration.Models;
using DYL.EmailIntegration.UI.Helpers;
using DYL.EmailIntegration.UI.Properties;
using log4net;

namespace DYL.EmailIntegration
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        [DllImport("user32.dll")]
        private static extern Boolean ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        const UInt32 SWP_NOSIZE = 0x0001;
        const UInt32 SWP_NOMOVE = 0x0002;
        const UInt32 SWP_SHOWWINDOW = 0x0040;

        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly LocalTimer _emailsTimer;
        private readonly LocalTimer _sessionTimer;
       
        public App()
        {
            if (CheckIfApplicationAlreadyRuning())
                return;

            log4net.Config.XmlConfigurator.Configure();
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += ExceptionEventHandler;
            _emailsTimer = new LocalTimer(Settings.Default.ServiceInterval, "Email");
           _sessionTimer = new LocalTimer(48 * 3600 * 1000, "Session");
        }

        private static bool CheckIfApplicationAlreadyRuning()
        {
            var currentProcess = Process.GetCurrentProcess();
            var runningProcess = (from process in Process.GetProcesses()
                where
                process.Id != currentProcess.Id &&
                process.ProcessName.Equals(
                    currentProcess.ProcessName,
                    StringComparison.Ordinal)
                select process).FirstOrDefault();

            if (runningProcess == null)
                return false;
          
            ShowWindow(runningProcess.MainWindowHandle, (int)ShowWindowCommands.SW_MAXIMIZE);
            SetWindowPos(runningProcess.MainWindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            return true;
        }

        static void ExceptionEventHandler(object sender, UnhandledExceptionEventArgs args)
        {
            Exception ex = (Exception)args.ExceptionObject;
            Log.Fatal("Unhandeled exception happened in application.", ex);
            MessageBox.Show("Unhandeled exception happened in application. See log file for details.", 
                "ERROR", MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            Log.Info("Application Started");
            Context.EmailQueue = new ObservableConcurrentQueue<Email>();
            HttpService.CreateService(Settings.Default.BaseUrl);
            ApplicationService.AutoLogin();
            _emailsTimer.Start(ApplicationService.EmailTimerEventHandler);
            _sessionTimer.Start(ApplicationService.SessionTimerEventHandler);
            base.OnStartup(e);
        }

       


        protected override void OnExit(ExitEventArgs e)
        {
            if (!Context.EmailQueue.IsEmpty)
            {
                Log.Info("Email queue is not empty.");
                MessageBox.Show("Email queue is not empty. Please review remaning emails.",
                    "WARNING", MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            _emailsTimer.Stop();
            _sessionTimer.Stop();
            Log.Info("Application Exited");
            base.OnExit(e);
        }
    }
}
