using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Contracts;
using DYL.EmailIntegration.Domain.Data;
using DYL.EmailIntegration.Service.Data;
using DYL.EmailIntegration.Service.IpcService;
using DYL.EmailIntegration.Service.Properties;
using log4net;
using Constants = DYL.EmailIntegration.Service.Helpers.Constants;

namespace DYL.EmailIntegration.Service
{
    public static class ApplicationService
    {
        private const string APP_ID = "DYL.NotificationService";
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static ServiceHost CreateNetPipeHost()
        {
            var address = Settings.Default.WcfNetPipeUrl;
            try
            {
                var serviceHost = new ServiceHost(typeof(LoginServer));
                var binding = new NetNamedPipeBinding(NetNamedPipeSecurityMode.Transport);
                serviceHost.AddServiceEndpoint(typeof(ILoginContract), binding, address);

                Log.Info($"WCF ServiceHost started. Address:{address}");

                return serviceHost;
            }
            catch (Exception ex)
            {
                Log.Error($"WCF ServiceHost failed to start. Address:{address}", ex);
            }
            return null;
        }

        public static async Task<string> GetNewSessionKey(Credentials credentials)
        {
            var credentialsHttpRequest = new CredentialsHttpRequest
            {
                password = credentials.password,
                email = credentials.email
            };
            return await HttpService.GetSessionKey(Constants.LoginUrl, credentialsHttpRequest)
                .ConfigureAwait(false);
        }

        public static async Task NotificationTimeHandler()
        {
            Log.Info("Notification timer elapsed event raised.");
            var total = await  GetNewEmailsTotal().ConfigureAwait(false);
            if (total != null && total > 0)
                ShowToastNotification((int)total);
        }

        public static async Task RenewSession()
        {
            Log.Info("Session timer elapsed event raised.");
           
            var credentials = Authentication.LoadCredentials(Constants.TokenFileName);
            if (credentials == null)
            {
                Log.Error("Failed to load credentials.");
                Context.Session = null;
                return;
            }

            var key =  await GetNewSessionKey(credentials).ConfigureAwait(false);

            if (string.IsNullOrEmpty(key))
            {
                Log.Error("Failed to get new session key.");
                Context.Session = null;
                return;
            }

            Context.Session = new Session(key, DateTime.Now);
        }

        private static async Task<int?> GetNewEmailsTotal()
        {
            var key = string.Empty;
            if (Context.Session == null)
            {
                Log.Info("Service session is null.");
                var credentials = Authentication.LoadCredentials(Constants.TokenFileName);
                if (credentials == null)
                {
                    Log.Error("Failed to load credentials.");
                    return null;
                }

                key = await GetNewSessionKey(credentials).ConfigureAwait(false);

                if (string.IsNullOrEmpty(key))
                {
                    Log.Error("Failed to get new session key.");
                    return null;
                }

                Context.Session = new Session(key, DateTime.Now);
            }
            else
            {
                key = Context.Session.Key;
            }
            var response = await HttpService.GetEmails(Constants.GetEmailsTotalUrl,
                new SessionKeyHttpRequest {session_key = key}).ConfigureAwait(false);
            var total = response?.Total;
            return total;
        }

        // Create and show the toast.
        private static void ShowToastNotification(int total)
        {
            try
            {
                // Get a toast XML template
                XmlDocument toastXml = ToastNotificationManager.GetTemplateContent(ToastTemplateType.ToastImageAndText02);

                // Fill in the text elements
                XmlNodeList stringElements = toastXml.GetElementsByTagName("text");
                stringElements[0].AppendChild(toastXml.CreateTextNode("DYL Email Adapter"));
                stringElements[1].AppendChild(toastXml.CreateTextNode($"There are {total} emails ready to be sent."));


                // Specify the absolute path to an image as a URI
                string imagePath = Path.GetFullPath("email_adapter_logo.png");
                XmlNodeList imageElements = toastXml.GetElementsByTagName("image");
                imageElements[0].Attributes.GetNamedItem("src").NodeValue = imagePath;

                // Create the toast and attach event listeners
                var toast = new ToastNotification(toastXml);
                toast.Activated += ToastActivated;
                toast.Dismissed += ToastDismissed;
                toast.Failed += ToastFailed;

                // Show the toast. Be sure to specify the AppUserModelId on your application's shortcut!
                ToastNotificationManager.CreateToastNotifier(APP_ID).Show(toast);
                Log.Info($"Show toast notification. There are {total} emails ready to be sent.");
            }
            catch (Exception ex)
            {
                Log.Error("Failed to show toast notification", ex);
            }
        }

        private static void ToastActivated(ToastNotification sender, object e)
        {
            Log.Info("The user activated the toast");
        }

        private static void ToastDismissed(ToastNotification sender, ToastDismissedEventArgs e)
        {
            string outputText = "";
            switch (e.Reason)
            {
                case ToastDismissalReason.ApplicationHidden:
                    outputText = "The app hid the toast using ToastNotifier.Hide";
                    break;
                case ToastDismissalReason.UserCanceled:
                    outputText = "The user dismissed the toast";
                    break;
                case ToastDismissalReason.TimedOut:
                    outputText = "The toast has timed out";
                    break;
            }
            Log.Info(outputText);
        }

        private static void ToastFailed(ToastNotification sender, ToastFailedEventArgs e)
        {
            Log.Info("The toast encountered an error.");
        }
    }
}
