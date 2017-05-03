using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain.Contracts;
using log4net;

namespace DYL.EmailIntegration.Service.IpcService
{
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.Single)]
    public class LoginServer : ILoginContract
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public bool Login(string username, string password)
        {
            Log.Debug($"Called Login Service.(Notification Windows Service)");
            return true;
        }
    }
}
