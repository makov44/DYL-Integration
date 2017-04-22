using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain.Data;
using log4net;
using log4net.Config;

namespace DYL.EmailIntegration.Domain
{
    public class HttpService
    {
        private static readonly HttpClient Client = new HttpClient();
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static void CreateService(string baseUrl)
        {
            Client.BaseAddress = new Uri(baseUrl);
            Client.DefaultRequestHeaders.Accept.Clear();
            Client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public static async Task<string> GetSessionKey(string requestUrl, Login login)
        {
            try
            {
                using (var response = await Client.PostAsJsonAsync(requestUrl, login))
                {
                    if (!response.IsSuccessStatusCode)
                        return String.Empty;

                    var result = await response.Content.ReadAsAsync<SessionKey>();
                    response.Content.Dispose();

                    return result?.session_key;
                }
            }
            catch (Exception ex)
            {
               Log.Error(ex);
               return String.Empty;
            }
        }

        public static async Task<EmailsHttpResponse> GetEmails(string requestUrl, SessionKey sessionkey)
        {
            try
            {
                using (var response = await Client.PostAsJsonAsync(requestUrl, sessionkey))
                {
                    if (!response.IsSuccessStatusCode)
                        return new EmailsHttpResponse();

                    var result = await response.Content.ReadAsAsync<EmailsHttpResponse>();
                    response.Content.Dispose();

                    return result;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                return new EmailsHttpResponse();
            }
        }
    }
}
