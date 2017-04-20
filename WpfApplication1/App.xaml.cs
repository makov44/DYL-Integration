using System;
using System.Collections.Concurrent;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Data;
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

        public App()
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += ExceptionEventHandler;
        }

        static void ExceptionEventHandler(object sender, UnhandledExceptionEventArgs args)
        {
            Exception ex = (Exception)args.ExceptionObject;
            Log.Fatal("Unhandeled exception happened in application.", ex);
            MessageBox.Show("Unhandeled exception happened in application. See log file for details.");
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            log4net.Config.XmlConfigurator.Configure();
            Log.Info("Application Started");
            Context.EmailQueue = new ConcurrentQueue<Email>();
            HttpService.CreateService(Settings.Default.BaseUrl);
            LocalTimer.Interval = Settings.Default.ServiceInterval;
            LocalTimer.Start(TimerEventHandler);
            base.OnStartup(e);
        }
        protected override void OnExit(ExitEventArgs e)
        {
            if (!Context.EmailQueue.IsEmpty)
            {
                Log.Info("Email queue is not empty.");
                MessageBox.Show("Email queue is not empty. Please review remaning emails.");
                return;
            }
            LocalTimer.Stop();
            Log.Info("Application Exited");
            base.OnExit(e);
        }

        private void TimerEventHandler()
        {
            Log.Info("Timer elapsed event raised.");
            if (string.IsNullOrEmpty(Context.SessionKey))
            {
                Log.Info("SessionKey is null or empty.");
                return;
            }

            if (!Context.EmailQueue.IsEmpty)
            {
                Log.Info("Email queue is not empty.");
                return;
            }

            var session = new Session
            {
                session_key = Context.SessionKey
            };
           
            System.Windows.Application.Current.Dispatcher.Invoke(async () =>
            {
                var response = await HttpService.GetEmails("api/outlook_emails", session);
                if (response == null || response.Count == 0  || response.Data == null)
                    return;
                response.Data.ForEach(x=> Context.EmailQueue.Enqueue(x));
            });
        }
    }
}
