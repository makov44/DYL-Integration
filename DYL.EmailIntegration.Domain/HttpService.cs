using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain.Contracts;
using DYL.EmailIntegration.Domain.Data;
using log4net;
using log4net.Config;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

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

        public static async Task<string> GetSessionKey(string requestUrl, CredentialsHttpRequest credentials)
        {
            try
            {
                using (var response = await Client.PostAsJsonAsync(requestUrl, credentials))
                {
                    var result = await response.Content.ReadAsAsync<SessionKey>();

                    if (!response.IsSuccessStatusCode)
                    {
                        Log.Error($"Failed to get new session key, statusCode:{response.StatusCode}, " +
                                  $"reason:{response.ReasonPhrase}, message:{result?.message}");
                        return string.Empty;
                    }
                   
                    response.Content.Dispose();

                    return result?.session_key;
                }
            }
            catch (Exception ex)
            {
               Log.Error(ex);
               return string.Empty;
            }
        }

        public static async Task<EmailsHttpResponse> GetEmails(string requestUrl, SessionKeyHttpRequest sessionkey)
        {
            try
            {
                using (var response = await Client.PostAsJsonAsync(requestUrl, sessionkey))
                {
                    var formatter = new JsonMediaTypeFormatter
                    {
                        SerializerSettings = new JsonSerializerSettings
                        {
                            Formatting = Newtonsoft.Json.Formatting.None,
                            ContractResolver = new CamelCasePropertyNamesContractResolver(),
                            DateTimeZoneHandling = DateTimeZoneHandling.Utc,
                            StringEscapeHandling = StringEscapeHandling.EscapeNonAscii,
                            MissingMemberHandling = MissingMemberHandling.Ignore,
                            Error = (sender, args) =>
                            {
                                args.ErrorContext.Handled = true;
                                Log.Error($"Error: {args.ErrorContext.Error} \r\n Member: {args.ErrorContext.Member}, " +
                                          $"Original Object: {args.ErrorContext.OriginalObject}, Path: {args.ErrorContext.Path} , Source: {sender}");
                            }
                        },
                    };

                    formatter.SupportedEncodings.Add(Encoding.GetEncoding("iso-8859-1"));
                    var result = await response.Content.ReadAsAsync<EmailsHttpResponse>(new[] { formatter});

                    if (!response.IsSuccessStatusCode)
                    {
                        Log.Error($"Failed to get emails, statusCode:{response.StatusCode}, " +
                                  $"reason:{response.ReasonPhrase}, message:{result?.Message}");
                        return new EmailsHttpResponse();
                    }

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

        public static async Task<BaseHttpResponse> PostStatus(string requestUrl, StatusHttpRequest status)
        {
            try
            {
                using (var response = await Client.PostAsJsonAsync(requestUrl, status).ConfigureAwait(false))
                {
                    var result = await response.Content.ReadAsAsync<BaseHttpResponse>().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        Log.Error($"Failed to post status, statusCode:{response.StatusCode}, " +
                                  $"reason:{response.ReasonPhrase}, message:{result?.Message}");
                        return new BaseHttpResponse();
                    }

                    response.Content.Dispose();

                    return result;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                return new BaseHttpResponse();
            }
        }
    }
}
