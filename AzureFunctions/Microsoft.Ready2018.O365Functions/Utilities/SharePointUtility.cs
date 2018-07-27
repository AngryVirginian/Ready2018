using AppForSharePointOnlineWebToolkit;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Ready2018.O365Functions.Models;
using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.IO;
using System.Security;
using System.Security.Cryptography.X509Certificates;


namespace Microsoft.Ready2018.O365Functions.Utilities
{
    public class SharePointUtility
    {
        private TraceWriter TraceWriter { get; set; }

        public SharePointUtility(TraceWriter traceWriter)
        {
            this.TraceWriter = traceWriter;            
        }

        #region Get ClientContext

        public ClientContext GetClientContext(SpoAuthenticationMethod method, string webUrl, ExecutionContext exeuctionContext)
        {
            switch (method)
            {
                case SpoAuthenticationMethod.UserNamePassword:
                    return (GetClientContextWithUserNamePassword(webUrl));
                case SpoAuthenticationMethod.SharePointAppIdentity:
                    return GetClientContextWithSharePointAppIdentity(webUrl);
                case SpoAuthenticationMethod.AzureAppIdentity:
                default:
                    return GetClientContextWithAzureAppIdentity(webUrl, exeuctionContext);
            }
        }

        public ClientContext GetClientContext(string webUrl, ExecutionContext exeuctionContext)
        {
            SpoAuthenticationMethod authMethod = this.GetSpoAuthenticationMethodFromAppSettings();
            return this.GetClientContext(authMethod, webUrl, exeuctionContext);
        }

        private ClientContext GetClientContextWithSharePointAppIdentity(string webUrl)
        {
            string appId = KeyVaultUtility.GetSecret(ConfigurationManager.AppSettings["sps:spGroupRequestAppIdKeyVaultSecretName"], this.TraceWriter).Result;
            string appSecret = KeyVaultUtility.GetSecret(ConfigurationManager.AppSettings["sps:spGroupRequestAppSecretKeyVaultSecretName"], this.TraceWriter).Result;
            this.TraceWriter.Info($"Creating SP Client Context with SharePoint App Identity to {webUrl} with appId { appId }");
            ClientContext clientContext = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(webUrl, appId, appSecret);
            return clientContext;
        }

        /// <summary>
        /// Get client context with Azure AD App and authenticate with certificate
        /// </summary>
        /// <param name="notification"></param>
        /// <param name="log"></param>
        /// <param name="executionContext"></param>
        /// <returns></returns>
        public ClientContext GetClientContextWithAzureAppIdentity(string webUrl, ExecutionContext executionContext)
        {
            //string url = String.Format("https://{0}{1}", ConfigurationManager.AppSettings["o365:SpoTenantUrl"], notification.SiteUrl);

            string spoTenantName = ConfigurationManager.AppSettings["o365:SpoTenantName"];

            string clientId = ConfigurationManager.AppSettings["aad:ApplicationId"];
            string clientSecret = ConfigurationManager.AppSettings["aad:ApplicationSecret"];
            string certName = ConfigurationManager.AppSettings["aad:ApplicationCertificatePrivateKeyFileName"];
            string certPassword = KeyVaultUtility.GetSecret(ConfigurationManager.AppSettings["kv:ApplicationCertificatePasswordSecretName"], this.TraceWriter).Result;
            //this.TraceWriter.Info($"Tenant Name is { spoTenantName } Client ID is {clientId } client Secret is { clientSecret } Cert name is { certName }  Cert password is { certPassword }");
            //Cert is at the root of the function
            //string certPath = Path.Combine(exeuctionContext.FunctionDirectory, certName);
            string certPath = Path.Combine(Directory.GetParent(executionContext.FunctionDirectory).FullName, certName);
            this.TraceWriter.Info($"Getting X509Certificate from { certPath }");
            this.TraceWriter.Info($"Parent path is {Path.Combine(Directory.GetParent(executionContext.FunctionDirectory).FullName, certName)}");
            X509Certificate2 cert = new X509Certificate2(certPath, certPassword);
            this.TraceWriter.Info($"Creating SP Client Context with Azure App Identity to {webUrl}");
            
            return new OfficeDevPnP.Core.AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(webUrl, clientId, spoTenantName, cert);

        }

        /// <summary>
        /// Get ClientContext with Azure Native App Identity with delegated access token
        /// </summary>
        /// <returns></returns>
        public ClientContext GetClientContextWithNativeAzureAppIdentity()
        {
            string resource = ConfigurationManager.AppSettings["o365:SpoTenantUrl"];
            string token =  GeneralUtility.GetAzureDelegatedAuthenticationToken(resource, this.TraceWriter).Result;

            return TokenHelper.GetClientContextWithAccessToken(resource, token);

        }


        /// <summary>
        /// Get SP Client Context based on user name and password stored in AppSettings
        /// </summary>
        public  ClientContext GetClientContextWithUserNamePassword(string webUrl)
        {
            string userName = KeyVaultUtility.GetSecret(ConfigurationManager.AppSettings["kv:ServicePrincipalNameSecretName"], this.TraceWriter).Result;
            string password = KeyVaultUtility.GetSecret(ConfigurationManager.AppSettings["kv:ServicePrincipalPasswordSecretName"], this.TraceWriter).Result;

            return GetClientContextWithUserNamePassword(webUrl, userName, password);

        }

        /// <summary>
        /// Get SP Client Context based on user name and password 
        /// </summary>
        public ClientContext GetClientContextWithUserNamePassword(string webUrl, string userName, string password)
        {
            SecureString securePassword = GeneralUtility.ConvertToSecureString(password); 
            return GetClientContextWithUserNamePassword(webUrl, userName, securePassword);

        }



        /// <summary>
        /// Get SP Client Context based on user name and password 
        /// </summary
        public ClientContext GetClientContextWithUserNamePassword(string url, string userName, SecureString password)
        {
            this.TraceWriter.Info($"Creating SPO Client Context with username and password to {url}");
            ClientContext clientContext = new ClientContext(url)
            {
                Credentials = new SharePointOnlineCredentials(userName, password)
            };

            return clientContext;
        }

        /// <summary>
        /// Get the SPO Authentication method from Function app settings
        /// </summary>
        /// <returns></returns>
        public SpoAuthenticationMethod GetSpoAuthenticationMethodFromAppSettings()
        {
            switch (ConfigurationManager.AppSettings["o365:SpoAuthenticationMethod"].ToLower())
            {
                case "usernamepassword":
                    return SpoAuthenticationMethod.UserNamePassword;
                case "sharepointappidentity":
                    return SpoAuthenticationMethod.SharePointAppIdentity;
                case "azureappidentity":
                    return SpoAuthenticationMethod.AzureAppIdentity;
                default:
                    throw new NotImplementedException($"Authentication method { ConfigurationManager.AppSettings["o365:SpoAuthenticationMethod"] } is not implemented");
            }
        }

        #endregion
    }
}
