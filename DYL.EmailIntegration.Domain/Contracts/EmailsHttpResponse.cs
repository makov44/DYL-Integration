using System.Collections.Generic;
using DYL.EmailIntegration.Domain.Data;

namespace DYL.EmailIntegration.Domain.Contracts
{
    public class EmailsHttpResponse : BaseHttpResponse
    {
        public int Total { get; set; }

        public int Page { get; set; }

        public List<Email> Data { get; set; }

        public int Limit { get; set; }
    }
}
