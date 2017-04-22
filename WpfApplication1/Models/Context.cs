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
        internal static Session Session { get; set; }
      
        internal static ObservableConcurrentBag<Email> EmailQueue { get; set; }

        internal static event EventHandler LoginClick;

        internal static event EventHandler OnLoginCompleted;

        internal static void RaiseLogInClickEvent(Login login)
        {
            LoginClick?.Invoke(new object(), new EventSessionArgs(login));
        }

        internal static void RaiseOnLoginCompletedEvent(LoginResult loginResult)
        {
            OnLoginCompleted?.Invoke(new object(), new EventLoginResultArg(loginResult));
        }
    }
}
