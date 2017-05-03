using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DYL.EmailIntegration.Service.Data
{
    public class Session
    {
        public string Key { get; }
        public DateTime TimeStamp { get; }

        public Session(string key, DateTime timeStamp)
        {
            Key = key;
            TimeStamp = timeStamp;
        }
    }
}
