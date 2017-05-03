using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Contracts;
using DYL.EmailIntegration.Service.Helpers;
using DYL.EmailIntegration.Service.IpcService;
using DYL.EmailIntegration.Service.Properties;
using log4net;

namespace DYL.EmailIntegration.Service
{
    public static class ApplicationService
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static void GetTotalEmails()
        {
                var credentials = new CredentialsHttpRequest
                {
                    password = "mfrager+joebob101@gmail.com",
                    email = "leads"
                };
            var key = HttpService.GetSessionKey(Constants.LoginUrl, credentials).Result;
            var session = new SessionKeyHttpRequest
            {
                session_key = key
            };
            var response = HttpService.GetEmails(Constants.GetEmailsTotalUrl, session).Result;
        }

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
    }
}
