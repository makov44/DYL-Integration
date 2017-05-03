using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DYL.EmailIntegration.Helpers
{
    public struct Constants
    {
        public const string TokenFileName = "token.txt";
        public const string GetEmailsUrl = "/api/sequence/emails";
        public const string LoginUrl = "/api/sequence/login";
        public const string StatusUrl = "/api/sequence/status";
        public const string LogoutUrl = "/api/sequence/logout";
    }
}
