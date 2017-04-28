using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DYL.EmailIntegration.Domain.Contracts
{
    public class CredentialsHttpRequest
    {
        public string email { get; set; }

        public string password { get; set; }
    }
}
