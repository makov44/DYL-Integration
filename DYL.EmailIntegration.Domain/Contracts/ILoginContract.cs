using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace DYL.EmailIntegration.Domain.Contracts
{
    [ServiceContract(Namespace = "http://dyl.com")]
    public interface ILoginContract
    {
        [OperationContract]
        bool Login(string username, string password);

        [OperationContract]
        bool Logout();
    }
}
