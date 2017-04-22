using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using log4net;

namespace DYL.EmailIntegration.Domain
{
    public class EncryptionService
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static readonly byte[] AditionalEntropy = { 6,5,78,45,67,78,23,25,76};

        public static byte[] Encrypt(string dataStr)
        {
            if(string.IsNullOrEmpty(dataStr))
                throw new Exception("Input parameter is null or empty.");
            try
            {
                var data = Encoding.UTF8.GetBytes(dataStr);
                // Encrypt the data using DataProtectionScope.CurrentUser. The result can be decrypted
                //  only by the same current user.
                return ProtectedData.Protect(data, AditionalEntropy, DataProtectionScope.CurrentUser);
            }
            catch (Exception ex)
            {
                Log.Error("Data was not encrypted. An error occurred.", ex);
                return new byte[0];
            }
        }

        public static string Decrypt(byte[] data)
        {
            try
            {
                //Decrypt the data using DataProtectionScope.CurrentUser.
                var result =  ProtectedData.Unprotect(data, AditionalEntropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(result);
            }
            catch (Exception ex)
            {
                Log.Error("Data was not decrypted. An error occurred.", ex);
                return string.Empty;
            }
        }
    }
}
