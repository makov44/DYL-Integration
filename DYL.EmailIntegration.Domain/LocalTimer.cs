using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using log4net;
using log4net.Config;

namespace DYL.EmailIntegration.Domain
{
    public static  class LocalTimer
    {
        private static readonly System.Timers.Timer Timer;
        public static int Interval { get; set; }
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static LocalTimer()
        {
            Timer = new Timer();
        }

        public static void Start(Action action)
        {
            if(Interval<=0)
                return;
            Timer.Interval = Interval;
            Timer.Elapsed += async (obj, e) => await Task.Factory.StartNew(action);
            Timer.AutoReset = true;
            Timer.Enabled = true;
            Timer.Start();
            Log.Info("Email service started.");
        }

        public static void Stop()
        {
            Timer.Enabled = false;
            Timer.Stop();
            Timer.Dispose();
            Log.Info("Email service stoped.");
        }
       
    }
}
