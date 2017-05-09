using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Contracts;
using DYL.EmailIntegration.Domain.Data;
using DYL.EmailIntegration.Service.Data;
using DYL.EmailIntegration.Service.Helpers;
using log4net;

namespace DYL.EmailIntegration.Service.IpcService
{
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.Single)]
    public class NetPipeService : ILoginContract, ISyncTimerContract
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public bool Login(string username, string password)
        {
            var credentials = new Credentials
            {
                password = password,
                email = username
            };

            Task.Run(async () =>
            {
                var key = await ApplicationService.GetNewSessionKey(credentials).ConfigureAwait(false);

                Context.Session = !string.IsNullOrEmpty(key) ? new Session(key, DateTime.Now) : null;

                if (string.IsNullOrEmpty(key))
                {
                    Log.Error("Failed to get new session key.");
                    return;
                }
                Authentication.SaveCredentials(credentials, Constants.TokenFileName);
            });
            Log.Info("Login to Notification service.");
            return true;
        }

        public bool Logout()
        {
            Context.Session = null;
            Authentication.CleanupCredentials(Constants.TokenFileName);
            Log.Info("Notification service was logout.");
            return true;
        }
        public void Notify()
        {
            Context.RaiseResetTimerEvent();
        }
    }
}
