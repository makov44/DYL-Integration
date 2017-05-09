using System.Collections.Generic;
using System.Runtime.Serialization;
using DYL.EmailIntegration.Domain.Data;

namespace DYL.EmailIntegration.Domain.Contracts
{
    public class EmailsHttpResponse : BaseHttpResponse
    {
        public int Count { get; set; }
        public int Total { get; set; }
        public List<Email> Data { get; set; }
    }
}
