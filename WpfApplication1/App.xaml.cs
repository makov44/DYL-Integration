using System;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Data;
using DYL.EmailIntegration.Helpers;
using DYL.EmailIntegration.Models;
using DYL.EmailIntegration.Properties;
using log4net;

namespace DYL.EmailIntegration
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly LocalTimer _emailsTimer;
        private readonly LocalTimer _sessionTimer;
       
        public App()
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += ExceptionEventHandler;
            _emailsTimer = new LocalTimer(Settings.Default.ServiceInterval, "Email");
           _sessionTimer = new LocalTimer(48 * 3600 * 1000, "Session");
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
            log4net.Config.XmlConfigurator.Configure();
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
