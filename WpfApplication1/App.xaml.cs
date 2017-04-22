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
            Context.EmailQueue = new ObservableConcurrentBag<Email>();
            Context.LoginClick += MainWindowLoginClick;
            HttpService.CreateService(Settings.Default.BaseUrl);
            LocalTimer.Interval = Settings.Default.ServiceInterval;
            LocalTimer.Start(TimerEventHandler);
            base.OnStartup(e);
        }

        private void MainWindowLoginClick(object sender, EventArgs e)
        {
            var eventSessionArgs = e as EventSessionArgs;

            if (eventSessionArgs == null)
                return;

             GetNewSessionKey(eventSessionArgs.Login, key =>
             {
                 Context.RaiseOnLoginCompletedEvent(string.IsNullOrEmpty(key)
                     ? LoginResult.Failed
                     : LoginResult.Success);

                 Context.Session = !string.IsNullOrEmpty(key) ? new Session(key, DateTime.Now) : null;

                 if (string.IsNullOrEmpty(key))
                     return;
                 var data = eventSessionArgs.Login.UserName + "|" + eventSessionArgs.Login.Password;
                 var dataEnrpt = EncryptionService.Encrypt(data);
                 if (dataEnrpt.Length == 0)
                     return;
                 IsolatedStorageService.Write(Constants.LoginFileName, dataEnrpt);
             });

           
        }

        private void GetNewSessionKey(Login login, Action<string> callback)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(async () =>
            {
                var key = await HttpService.GetSessionKey("api/outlook", login);
                callback(key);
            });
        }

        protected override void OnExit(ExitEventArgs e)
        {
            if (!Context.EmailQueue.IsEmpty)
            {
                Log.Info("Email queue is not empty.");
                MessageBox.Show("Email queue is not empty. Please review remaning emails.");
            }
            LocalTimer.Stop();
            Log.Info("Application Exited");
            base.OnExit(e);
        }

        private void TimerEventHandler()
        {
            Log.Info("Timer elapsed event raised.");
            if (Context.Session == null)
            {
                Log.Info("Session is null");
                return;
            }

            if (!Context.EmailQueue.IsEmpty)
            {
                Log.Info("Email queue is not empty.");
                return;
            }

            if (Context.Session.TimeStamp.AddHours(48) < DateTime.Now)
            {
                var dataEnrpt = IsolatedStorageService.Read(Constants.LoginFileName);
                var data = EncryptionService.Decrypt(dataEnrpt);
                if (string.IsNullOrEmpty(data))
                {
                    Log.Error("Decrypted data is not valid");
                    return;
                }
                    
                var array = data.Split('|');

                if (array.Length != 2)
                {
                    Log.Error("Decrypted data is not valid");
                    return;
                }

                var login = new Login
                {
                    UserName = array[0],
                    Password = array[1]
                };
                
                GetNewSessionKey(login, key =>
                {
                    Context.Session = !string.IsNullOrEmpty(key) 
                    ? new Session(key, DateTime.Now) 
                    : null;

                    if (Context.Session != null)
                        GetEmails(Context.Session.Key); 
                });
            }
            else
            {
                GetEmails(Context.Session.Key);
            }
        }

        private  void GetEmails(string key)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(async () =>
            {
                try
                {
                    var session = new SessionKey
                    {
                        session_key = key
                    };
                    var response = await HttpService.GetEmails("api/outlook_emails", session);
                    if (response == null || response.Count == 0 || response.Data == null)
                        return;
                    response.Data.ForEach(x => Context.EmailQueue.Enqueue(x));

                    if (Context.EmailQueue.Count > 1000)
                        Log.Error("Email Queue has more then 1000 items.");
                }
                catch (Exception ex)
                {
                    Log.Error(ex);
                }
            });
        }
    }
}
