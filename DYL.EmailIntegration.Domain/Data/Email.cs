using System.Runtime.Serialization;

namespace DYL.EmailIntegration.Domain.Data
{
    public class Email
    {
        public string Id { get; set; }
       
        public string To { get; set; }
      
        public string Subject { get; set; }
      
        public string Body { get; set; }       
    }
}
