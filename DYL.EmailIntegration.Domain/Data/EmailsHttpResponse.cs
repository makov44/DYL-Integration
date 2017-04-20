using System.Collections.Generic;

namespace DYL.EmailIntegration.Domain.Data
{
    public class EmailsHttpResponse
    {
        public int Count { get; set; }

        public int Page { get; set; }

        public List<Email> Data { get; set; }

        public int Limit { get; set; }
    }
}
