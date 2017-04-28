using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Contracts;
using DYL.EmailIntegration.Service.Helpers;

namespace DYL.EmailIntegration.Service
{
    public static class ApplicationService
    {
        public static void GetTotalEmails()
        {
                var credentials = new CredentialsHttpRequest
                {
                    password = "mfrager+joebob101@gmail.com",
                    email = "leads"
                };
            var key = HttpService.GetSessionKey(Constants.LoginUrl, credentials).Result;
            var session = new SessionKeyHttpRequest
            {
                session_key = key
            };
            var response = HttpService.GetEmails(Constants.GetEmailsTotalUrl, session).Result;
        }
    }
}
