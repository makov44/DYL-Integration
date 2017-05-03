using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using log4net;

namespace DYL.EmailIntegration.Service
{
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(INotificationActivationCallback))]
    [Guid("23A5B06E-20BB-4E7E-A0AC-6982ED6A6041"), ComVisible(true)]
    public class NotificationActivator : INotificationActivationCallback
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public void Activate(string appUserModelId, string invokedArgs, NOTIFICATION_USER_INPUT_DATA[] data, uint dataCount)
        {
            Log.Info("Toast activated.(NotificationActivator)");
        }

        public static void Initialize()
        {
            Log.Info("NotificationActivator was initialized.");
            regService = new RegistrationServices();

            cookie = regService.RegisterTypeForComClients(
                typeof(NotificationActivator),
                RegistrationClassContext.LocalServer,
                RegistrationConnectionType.MultipleUse);
        }
        public static void Uninitialize()
        {
            if (cookie != -1)
            {
                regService?.UnregisterTypeForComClients(cookie);
            }
        }

        private static int cookie = -1;
        private static RegistrationServices regService = null;
    }
}
