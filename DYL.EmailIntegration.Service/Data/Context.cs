using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DYL.EmailIntegration.Service.Data
{
    internal static class Context
    {
        public static event EventHandler ResetTimer;
        internal static Session Session { get; set; }

        public static void RaiseResetTimerEvent()
        {
            ResetTimer?.Invoke(new object(), new EventArgs());
        }
    }
}
