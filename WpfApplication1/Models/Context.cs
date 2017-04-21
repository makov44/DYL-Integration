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
        internal static string SessionKey { get; set; }
      
        internal static ObservableConcurrentBag<Email> EmailQueue { get; set; }
      
    }
}
