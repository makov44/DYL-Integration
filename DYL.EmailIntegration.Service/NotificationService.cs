using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceModel;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Contracts;
using DYL.EmailIntegration.Service.Helpers;
using DYL.EmailIntegration.Service.Properties;
using log4net;

namespace DYL.EmailIntegration.Service
{
    public partial class NotificationService : ServiceBase
    {
        private const string APP_ID = "DYL.NotificationService";
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly LocalTimer _emailsTimer;
        private readonly ServiceHost _serviceHost;
        public NotificationService()
        {
            _emailsTimer = new LocalTimer(Settings.Default.NotificationInterval, "Notification");
            _serviceHost = ApplicationService.CreateNetPipeHost();
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Log.Info("Notification service started");
            if (string.IsNullOrEmpty(Settings.Default.BaseUrl))
            {
                Log.Error("Settings property 'BaseUrl' in config file is empty.");
                return;
            }
            HttpService.CreateService(Settings.Default.BaseUrl);
            NotificationActivator.Initialize();
            try
            {
                _serviceHost?.Open();
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                _serviceHost?.Abort();
            }
           
            _emailsTimer.Start(NotificationTimeHandler);
          //  _sessionTimer.Start(ApplicationService.SessionTimerEventHandler);

        }

        private void NotificationTimeHandler()
        {
            Log.Info("Notification timer elapsed event raised.");
            var total = GetNewEmailsTotal();
            if (total != null && total > 0)
                ShowToastNotification((int)total);
        }

        protected override void OnStop()
        {
            NotificationActivator.Uninitialize();
            _serviceHost?.Close();
            _emailsTimer.Stop();
            Log.Info("Notification service Exited");
        }

        private int? GetNewEmailsTotal()
        {
            var credentialsHttpRequest = new CredentialsHttpRequest
            {
                password = "leads",
                email = "65654@sam.com"
            };
            var key = HttpService.GetSessionKey(Constants.LoginUrl, credentialsHttpRequest).Result;
            var response = HttpService.GetEmails(Constants.GetEmailsTotalUrl, new SessionKeyHttpRequest { session_key = key}).Result;
            var total = response?.Total;
            return total;
        }

        // Create and show the toast.
        private void ShowToastNotification(int total)
        {
            try
            {
                // Get a toast XML template
                XmlDocument toastXml = ToastNotificationManager.GetTemplateContent(ToastTemplateType.ToastImageAndText02);

                // Fill in the text elements
                XmlNodeList stringElements = toastXml.GetElementsByTagName("text");
                stringElements[0].AppendChild(toastXml.CreateTextNode("DYL"));
                stringElements[1].AppendChild(toastXml.CreateTextNode($"There are {total} new emails"));


                // Specify the absolute path to an image as a URI
                String imagePath = Path.GetFullPath("logo.png");
                XmlNodeList imageElements = toastXml.GetElementsByTagName("image");
                imageElements[0].Attributes.GetNamedItem("src").NodeValue = imagePath;

                // Create the toast and attach event listeners
                var toast = new ToastNotification(toastXml);
                toast.Activated += ToastActivated;
                toast.Dismissed += ToastDismissed;
                toast.Failed += ToastFailed;

                // Show the toast. Be sure to specify the AppUserModelId on your application's shortcut!
                ToastNotificationManager.CreateToastNotifier(APP_ID).Show(toast);
                Log.Info($"Show toast notification. There are {total} new emails");
            }
            catch (Exception ex)
            {
                Log.Error("Failed to show toast notification", ex);
            }
        }

        private void ToastActivated(ToastNotification sender, object e)
        {
            Log.Info("The user activated the toast");
        }

        private void ToastDismissed(ToastNotification sender, ToastDismissedEventArgs e)
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

        private void ToastFailed(ToastNotification sender, ToastFailedEventArgs e)
        {
            Log.Info("The toast encountered an error.");
        }
    }
}
