using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceModel;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Contracts;
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
        /// <summary> 
        /// Mutex used to allow only one instance. 
        /// </summary> 
        private Mutex _mutex;

        /// <summary> 
        /// Name of mutex to use. Should be unique for all applications. 
        /// </summary> 
        public const string MutexName = "DYL.EmailIntegration.Mutex";

        /// <summary> 
        /// Sets the foreground window. 
        /// </summary> 
        /// <param name="hWnd">Window handle to bring to front.</param> 
        /// <returns></returns> 
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private LocalTimer _emailsTimer;
        private LocalTimer _sessionTimer;
       
        public App()
        {
            TryToStartApplication(SetupApp);
        }

        private void SetupApp()
        {
            log4net.Config.XmlConfigurator.Configure();
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += ExceptionEventHandler;
            _emailsTimer = new LocalTimer(Settings.Default.ServiceInterval, "Email");
            _sessionTimer = new LocalTimer(Settings.Default.SessionExpirationInterval, "Session");
        }

      

        private void TryToStartApplication(Action callback)
        {
            // Try to grab mutex 
            bool createdNew;
            _mutex = new Mutex(true, MutexName, out createdNew);

            if (!createdNew)
            {
                // Bring other instance to front and exit. 
                var current = Process.GetCurrentProcess();
                foreach (var process in Process.GetProcessesByName(current.ProcessName))
                {
                    if (process.Id != current.Id)
                    {
                        SetForegroundWindow(process.MainWindowHandle);
                        break;
                    }
                }
                Application.Current.Shutdown();
            }
            else
            {
                callback();
                Exit += CloseMutexHandler;
            }
        }

        /// <summary> 
        /// Handler that closes the mutex. 
        /// </summary> 
        /// <param name="sender">Sender of the event.</param> 
        /// <param name="e">Event arguments.</param> 
        protected virtual void CloseMutexHandler(object sender, EventArgs e)
        {
            _mutex?.Close();
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
            ApplicationService.SyncTimerService();
            _emailsTimer.Start(ApplicationService.EmailTimerEventHandler);
            _sessionTimer.Start(ApplicationService.RenewSession);
            base.OnStartup(e);
        }

       


        protected override void OnExit(ExitEventArgs e)
        {
            _emailsTimer?.Stop();
            _sessionTimer?.Stop();
            _emailsTimer?.Dispose();
            _sessionTimer?.Dispose();
            Log?.Info("Application Exited");
            base.OnExit(e);
        }
    }
}
