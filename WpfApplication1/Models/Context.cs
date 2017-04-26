using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain.Data;
using DYL.EmailIntegration.Helpers;

namespace DYL.EmailIntegration.Models
{
    internal static class Context
    {

        internal static event PropertyChangedEventHandler PropertyChanged;
        private static readonly object Locker = new object();

        private static Session _session;
        internal static Session Session
        {
            get { return _session; }
            set
            {
                lock (Locker)
                {
                    if (_session == value)
                        return;
                    _session = value;
                    PropertyChanged?.Invoke(new object(), new PropertyChangedEventArgs("Session"));
                }
            }
        }
      
        internal static ObservableConcurrentQueue<Email> EmailQueue { get; set; }
    }
}
