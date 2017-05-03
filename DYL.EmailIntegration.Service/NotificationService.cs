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
using DYL.EmailIntegration.Service.Data;
using DYL.EmailIntegration.Service.Helpers;
using DYL.EmailIntegration.Service.Properties;
using log4net;

namespace DYL.EmailIntegration.Service
{
    public partial class NotificationService : ServiceBase
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly LocalTimer _emailsTimer;
        private readonly LocalTimer _sessionTimer;
        private readonly ServiceHost _serviceHost;
        public NotificationService()
        {
            _emailsTimer = new LocalTimer(Settings.Default.NotificationInterval, "Notification");
            _sessionTimer = new LocalTimer(Settings.Default.SessionExpirationInterval, "Session");
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
           
            _emailsTimer.StartTask(ApplicationService.NotificationTimeHandler);
            _sessionTimer.StartTask(ApplicationService.RenewSession);
        }

       

        protected override void OnStop()
        {
            NotificationActivator.Uninitialize();
            _serviceHost?.Close();
            _emailsTimer.Stop();
            Log.Info("Notification service Exited");
        }
    }
}
