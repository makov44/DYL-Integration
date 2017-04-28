using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Contracts;
using DYL.EmailIntegration.Service.Helpers;
using log4net;

namespace DYL.EmailIntegration.Service
{
    public partial class NotificationService : ServiceBase
    {
        private const string APP_ID = "DYL.NotificationService";
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly LocalTimer _emailsTimer;
        public NotificationService()
        {
            _emailsTimer = new LocalTimer(30000, "Notification");
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Log.Info("Sync service started");
            HttpService.CreateService("http://dyl1-cy.getdyl.com");
            NotificationActivator.Initialize();
            _emailsTimer.Start(NotificationTimeHandler);
            //    _sessionTimer.Start(ApplicationService.SessionTimerEventHandler);

        }

        private void NotificationTimeHandler()
        {
            var total = GetNewEmailsTotal();
            if (total != null && total > 0)
                ShowToastNotification((int)total);
        }

        protected override void OnStop()
        {
            _emailsTimer.Stop();
            NotificationActivator.Uninitialize();
            Log.Info("Sync service Exited");
        }


        private int? GetNewEmailsTotal()
        {
            var credentialsHttpRequest = new CredentialsHttpRequest
            {
                password = "leads",
                email = "mfrager+joebob101@gmail.com"
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
                String imagePath = new System.Uri(Path.GetFullPath("logo.png")).AbsoluteUri;
                XmlNodeList imageElements = toastXml.GetElementsByTagName("image");
                imageElements[0].Attributes.GetNamedItem("src").NodeValue = imagePath;

                // Create the toast and attach event listeners
                var toast = new ToastNotification(toastXml);
                toast.Activated += ToastActivated;
                toast.Dismissed += ToastDismissed;
                toast.Failed += ToastFailed;

                // Show the toast. Be sure to specify the AppUserModelId on your application's shortcut!
                ToastNotificationManager.CreateToastNotifier(APP_ID).Show(toast);
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
