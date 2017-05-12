using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Contracts;
using DYL.EmailIntegration.Domain.Data;
using DYL.EmailIntegration.Models;
using DYL.EmailIntegration.UI.Helpers;
using DYL.EmailIntegration.UI.Properties;
using log4net;

namespace DYL.EmailIntegration.Helpers
{
    public static class ApplicationService
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static void GetEmails(string key)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(async () =>
            {
                try
                {
                    var session = new SessionKeyHttpRequest
                    {
                        session_key = key
                    };
                    var response = await HttpService.GetEmails(Constants.GetEmailsUrl, session);
                    if (response == null || response.Count == 0 || response.Data == null)
                        return;

                    response.Data.ForEach(x =>
                    {
                        if(!Context.EmailQueue.ContainsId(x))
                            Context.EmailQueue.Enqueue(x);
                    });
                }
                catch (Exception ex)
                {
                    Log.Error(ex);
                }
            });
        }

        public static void GetNewSessionKey(Credentials credentials, Action<string> callback)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(async () =>
            {
                var credentialsHttpRequest = new CredentialsHttpRequest
                {
                    password = credentials.password,
                    email = credentials.email
                };
                var key = await HttpService.GetSessionKey(Constants.LoginUrl, credentialsHttpRequest);
                callback(key);
            });
        }

        public static void AutoLogin()
        {
            var credentials = Authentication.LoadCredentials(Constants.TokenFileName);
            if (credentials != null)
            {
                GetNewSessionKey(credentials, key =>
                {
                    Context.Session = !string.IsNullOrEmpty(key)
                        ? new Session(key, DateTime.Now)
                        : null;
                });
            }
        }

        public static void EmailTimerEventHandler()
        {
            Log.Info("Email timer elapsed event raised.");
            if (Context.Session == null)
            {
                Log.Info("Session is null");
                return;
            }

            GetEmails(Context.Session.Key);
        }

        public static void RenewSession()
        {
            Log.Info("Session timer elapsed event raised.");

            var credentials = Authentication.LoadCredentials(Constants.TokenFileName);
            if (credentials == null)
            {
                Log.Error("Failed to load credentials.");
                System.Windows.Application.Current.Dispatcher.Invoke(() => { Context.Session = null; });
                return;
            }

            GetNewSessionKey(credentials, key =>
            {
                Context.Session = !string.IsNullOrEmpty(key)
                    ? new Session(key, DateTime.Now)
                    : null;
            });
        }

        public static void PostEmailStatus(Status status)
        {
            var statusHttpRequest = new StatusHttpRequest
            {
                id = status.Id,
                sequence_id = status.SequenceId,
                status = status.StatusName,
                session_key = Context.Session.Key
            };
            Task.Run(() => 
                  HttpService.PostStatus(Constants.StatusUrl, statusHttpRequest));
        }

        private static ChannelFactory<T> CreateChannelFactory<T>()
        {
            var fullAddress = Settings.Default.WcfNetPipeUrl;
            var binding = new NetNamedPipeBinding(NetNamedPipeSecurityMode.Transport);
            var ep = new EndpointAddress(fullAddress);
            var factory = new ChannelFactory<T>(binding, ep);
            return factory;
        }

        public static void AutoLoginNotificationService(Credentials credentials)
        {
            var factory = CreateChannelFactory<ILoginContract>();
            try
            {
                var channel = factory.CreateChannel();
                var isSuccess = channel.Login(credentials.email, credentials.password);

                if(!isSuccess) 
                    Log.Error("Faled to call auto login service. (Notification Service)");
                
                factory.Close();
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                factory.Abort();
            }
        }

        public static void SyncTimerService()
        {
            var factory = CreateChannelFactory<ISyncTimerContract>();
            try
            {
                var channel = factory.CreateChannel();
                channel.Notify();
                factory.Close();
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                factory.Abort();
            }
        }

        public static void LogoutNotificationService()
        {
            var factory = CreateChannelFactory<ILoginContract>();
            try
            {
                var channel = factory.CreateChannel();
                var isSuccess = channel.Logout();

                if (!isSuccess)
                    Log.Error("Faled to call logout service. (Notification Service)");

                factory.Close();
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                factory.Abort();
            }
        }
    }
}
