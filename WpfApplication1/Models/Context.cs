using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain.Data;

namespace DYL.EmailIntegration.Models
{
    internal static class Context
    {
        internal static string SessionKey { get; set; }
        internal static ConcurrentQueue<Email> EmailQueue { get; set; }
    }
}
