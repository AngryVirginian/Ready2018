using System;
using System.Threading.Tasks;
using Microsoft.Azure.KeyVault;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Configuration;
using Microsoft.Azure.WebJobs.Host;

namespace Microsoft.Ready2018.O365Functions.Utilities
{
    public class KeyVaultUtility
    {
        public static async Task<string> GetSecret(string secretName, TraceWriter log)
        {
            var kv = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(GetKeyVaultAccessToken));
            //log.Info($"Getting secret for {secretName}");
            var sec = await kv.GetSecretAsync(ConfigurationManager.AppSettings["kv:KeyVaultUrl"], secretName);
            return sec.Value;
        }

        //the method that will be provided to the KeyVaultClient
        private static async Task<string> GetKeyVaultAccessToken(string authority, string resource, string scope)
        {
            ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["aad:ApplicationId"],
                            ConfigurationManager.AppSettings["aad:ApplicationSecret"]);

            try
            {
                var authContext = new AuthenticationContext(authority);

                AuthenticationResult result = await authContext.AcquireTokenAsync(resource, clientCred);
                if (result == null)
                    throw new InvalidOperationException("Failed to obtain the JWT token");
                return result.AccessToken;

            }
            catch (Exception e)
            {
                throw new Exception($"Error getting keyvault authentication token for {clientCred.ClientId}", e);
            }

        }
    }
}
