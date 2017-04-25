using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Data;
using log4net;

namespace DYL.EmailIntegration.Helpers
{
    public static class Authentication
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static Credentials LoadCredentials()
        {
            try
            {
                var dataEnrpt = IsolatedStorageService.Read(Constants.LoginFileName);

                if (dataEnrpt.Length == 0)
                {
                    Log.Error("Failed to read credentials from isolated storage.");
                    return null;
                }

                var data = EncryptionService.Decrypt(dataEnrpt);
                if (string.IsNullOrEmpty(data))
                {
                    Log.Error("Decrypted data is not valid");
                    return null;
                }

                var array = data.Split('|');

                if (array.Length != 2)
                {
                    Log.Error("Decrypted data is not valid");
                    return null;
                }

                var credentials = new Credentials
                {
                    email = array[0],
                    password = array[1]
                };
                return credentials;
            }
            catch (Exception ex)
            {
                Log.Error("Failed to load credentials.", ex);
                return null;
            }
        }

        public static void SaveCredentials(Credentials credentials)
        {
            try
            {
                var data = credentials.email + "|" + credentials.password;
                var dataEnrpt = EncryptionService.Encrypt(data);
                if (dataEnrpt.Length == 0)
                    return;
                IsolatedStorageService.Write(Constants.LoginFileName, dataEnrpt);
            }
            catch (Exception ex)
            {
                Log.Error("Failed to save credentials.", ex);
            }
        }
    }
}
