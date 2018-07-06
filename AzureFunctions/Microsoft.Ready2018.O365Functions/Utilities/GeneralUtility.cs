using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using Microsoft.Ready2018.O365Functions.Models;
using System.Configuration;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace Microsoft.Ready2018.O365Functions.Utilities
{
    public class GeneralUtility
    {
        /// <summary>
        /// Get SecureString from a plain string
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static SecureString ConvertToSecureString(string input)
        {
            SecureString secureString = new SecureString();
            foreach (char c in input.ToCharArray())
            {
                secureString.AppendChar(c);
            }
            return secureString;
        }

        /// <summary>
        /// Get Delegated access token to AAD with credentials stored in AppSettings
        /// </summary>
        /// <returns></returns>
        public static async Task<string> GetAzureDelegatedAuthenticationToken(string resource, TraceWriter log)
        {
            log.Info($"Getting delegated app access token");

            string authority = $"https://login.microsoftonline.com/{ ConfigurationManager.AppSettings["o365:SpoTenantName"] }";
            string tenantId = ConfigurationManager.AppSettings["aad:TenantId"];
            string clientId = ConfigurationManager.AppSettings["aad:NativeAppId"];

            UserCredential userCredential = new UserCredential(
                KeyVaultUtility.GetSecret(ConfigurationManager.AppSettings["kv:ServicePrincipalNameSecretName"], log).Result,
                KeyVaultUtility.GetSecret(ConfigurationManager.AppSettings["kv:ServicePrincipalPasswordSecretName"], log).Result);

            AuthenticationContext context = new AuthenticationContext(authority);
            var result = await context.AcquireTokenAsync(resource, clientId, userCredential);
            log.Info($"auth result for { userCredential.UserName } is { result.UserInfo } { result.AccessTokenType }");

            return result.AccessToken;
        }


    }
}
