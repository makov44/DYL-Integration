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
    public class LocalTimer
    {
        private readonly System.Timers.Timer _timer;

        private readonly string _name;

        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public  LocalTimer(int interval, string name)
        {
            _name = name;
            _timer = new Timer(interval);
        }

        public void Start(Action action)
        {
            if(_timer.Interval <= 0)
                return;
            _timer.Elapsed += (obj, e) => Task.Factory.StartNew(action);
            _timer.AutoReset = true;
            _timer.Enabled = true;
            _timer.Start();
            Log.Info($"{_name} timer started.");
        }

        public void StartTask(Func<Task> action)
        {
            if (_timer.Interval <= 0)
                return;
            _timer.Elapsed +=  (obj, e) => action();
            _timer.AutoReset = true;
            _timer.Enabled = true;
            _timer.Start();
            Log.Info($"{_name} timer started.");
        }

        public void Stop()
        {
            _timer.Enabled = false;
            _timer.Stop();
            _timer.Dispose();
            Log.Info($"{_name} timer stoped.");
        }
    }
}
