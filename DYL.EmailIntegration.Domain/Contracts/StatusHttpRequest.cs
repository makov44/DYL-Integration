namespace DYL.EmailIntegration.Domain.Contracts
{
    public class StatusHttpRequest
    {
        public string session_key { get; set; }

        public string id { get; set; }

        public string status { get; set; }
    }
}
