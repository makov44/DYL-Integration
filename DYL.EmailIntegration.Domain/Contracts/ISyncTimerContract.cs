using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace DYL.EmailIntegration.Domain.Contracts
{
    [ServiceContract(Namespace = "http://dyl.com")]
    public interface ISyncTimerContract
    {
        [OperationContract]
        void Notify();
    }
}
