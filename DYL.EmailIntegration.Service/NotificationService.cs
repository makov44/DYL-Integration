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
        private LocalTimer _emailsTimer;
        private LocalTimer _sessionTimer;
        private ServiceHost _serviceHost;
        public NotificationService()
        {
            InitializeComponent();
        }

        private void Context_ResetTimer(object sender, EventArgs e)
        {
            _emailsTimer.Reset();
        }

        protected override void OnStart(string[] args)
        {
            Log.Info("Notification service started");
            if (string.IsNullOrEmpty(Settings.Default.BaseUrl))
            {
                Log.Error("Settings property 'BaseUrl' in config file is empty.");
                return;
            }

            _serviceHost = ApplicationService.CreateNetPipeHost();
            _emailsTimer = new LocalTimer(Settings.Default.NotificationInterval * 60000, "Notification");
            _sessionTimer = new LocalTimer(Settings.Default.SessionExpirationInterval * 60000, "Session");
            HttpService.CreateService(Settings.Default.BaseUrl);
            Context.ResetTimer += Context_ResetTimer;

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
            _emailsTimer?.Stop();
            _sessionTimer?.Stop();
            Context.ResetTimer -= Context_ResetTimer;
            NotificationActivator.Uninitialize();
            _emailsTimer?.Dispose();
            _sessionTimer?.Dispose();
            _serviceHost?.Close();
            
            Log.Info("Notification Service Stoped");
        }
    }
}
