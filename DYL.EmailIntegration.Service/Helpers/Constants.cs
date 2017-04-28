using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DYL.EmailIntegration.Service.Helpers
{
    public struct Constants
    {
        public const string GetEmailsTotalUrl = "api/sequence/emails?count_only=1";
        public const string LoginUrl = "/api/sequence/login";
        public const string LogoutUrl = "/api/sequence/logout";
    }
}
