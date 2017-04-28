using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Contracts;
using DYL.EmailIntegration.Domain.Data;
using DYL.EmailIntegration.Models;
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
                    if (response == null || response.Total == 0 || response.Data == null)
                        return;
                    response.Data.ForEach(x => Context.EmailQueue.Enqueue(x));

                    if (Context.EmailQueue.Count > 200)
                        Log.Error("Email Queue has more then 1000 items.");
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
            var credentials = Authentication.LoadCredentials();
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

            if (!Context.EmailQueue.IsEmpty)
            {
                Log.Info("Email queue is not empty.");
                return;
            }
          
            GetEmails(Context.Session.Key);
        }

        public static void SessionTimerEventHandler()
        {
            Log.Info("Session timer elapsed event raised.");
            var credentials = Authentication.LoadCredentials();
            if (credentials == null)
            {
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
                id = status.EmailId,
                status = status.StatusName,
                session_key = Context.Session.Key
            };
            Task.Run(async () => 
                 await HttpService.PostStatus(Constants.StatusUrl, statusHttpRequest));
        }
    }
}
