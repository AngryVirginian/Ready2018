using Microsoft.Azure.WebJobs.Host;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using Microsoft.Ready2018.O365Functions.Models;

namespace Microsoft.Ready2018.O365Functions.Utilities
{
    public class GraphApiUtility
    {
        private TraceWriter TraceWriter { get; set; }

        public GraphApiUtility(TraceWriter log)
        {
            this.TraceWriter = log;
        }


        #region Authentication Token

        /// <summary>
        /// Get authentication token using Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext
        /// </summary>
        /// <param name="resource"></param>
        /// <returns></returns>
        public async Task<string> GetGraphApiAuthenticationToken()
        {
            string tenantId = ConfigurationManager.AppSettings["aad:TenantId"];
            string authority = $"https://login.windows.net/{ tenantId }/oauth2/token";
            string clientId = ConfigurationManager.AppSettings["aad:ApplicationId"];
            string clientSecret = ConfigurationManager.AppSettings["aad:ApplicationSecret"];
            string resource = ConfigurationManager.AppSettings["gph:GraphApiUrl"];

            this.TraceWriter.Info($"Getting authentication token for {resource}");

            AuthenticationContext authContext = new AuthenticationContext(authority);

            ClientCredential clientCredential = new ClientCredential(clientId, clientSecret);
            AuthenticationResult authResult = await authContext.AcquireTokenAsync(ConfigurationManager.AppSettings["gph:GraphApiTokenResourceUrl"], clientCredential);

            //this.TraceWriter.Info($"Authentication token is {authResult.AccessToken}");
            return authResult.AccessToken;
        }

        /// <summary>
        /// Get Delegated access token to AAD with credentials stored in AppSettings
        /// </summary>
        /// <returns></returns>
        public async Task<string> GetGraphApiDelegatedAuthenticationToken()
        {
            
            string resource = ConfigurationManager.AppSettings["gph:GraphApiTokenResourceUrl"];
            return await GeneralUtility.GetAzureDelegatedAuthenticationToken(resource, this.TraceWriter);

        }

        

        #endregion

        /// <summary>
        /// Whether a user with the nickname parameter already exists.
        /// </summary>
        public async Task<bool> DoesUserEmailNicknameExists(string emailNickname)
        {
            using (HttpClient client = new HttpClient())
            {
                //Get authentication token
                string token = await this.GetGraphApiAuthenticationToken();
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                //Check both whether the group already exists or whether a user with the same nickname already exist
                return DoesGroupExists(client, emailNickname) | DoesUserEmailNicknameExists(client, emailNickname);
            }
        }


        // <summary>
        // Whether a user with the nickname parameter already exists.
        // </summary>
        private bool DoesUserEmailNicknameExists(HttpClient client, string emailNickname)
        {
            //https://graph.microsoft.com/v1.0/users?$filter=mailNickName eq 'LEGAL'

            var response = client.GetStringAsync($"{ ConfigurationManager.AppSettings["gph:GraphApiUrl"] }/users?$filter=mailNickName eq '{ emailNickname }'").Result;
            var responseData = JsonConvert.DeserializeObject<GraphApiGetUserResponse>(response);

            if (responseData != null && responseData.GraphUserObjects != null && responseData.GraphUserObjects.Count > 0)
            {
                //If there is more than zero user in response
                return true;
            }
            else
            {
                return false;
            }

        }


        /// <summary>
        /// Whether a unified, security, or distribution group with the nickname parameter already exists.
        /// </summary>
        /// <param name="client"></param>
        /// <param name="mailNickname"></param>
        /// <returns></returns>
        private bool DoesGroupExists(HttpClient client, string mailNickname)
        {
            //"https://graph.microsoft.com/v1.0/groups?`$filter=startswith(mail,'$EMail')"
            //https://graph.microsoft.com/v1.0/groups?$filter=mailNickName eq 'Legal'

            var response = client.GetStringAsync($"{ ConfigurationManager.AppSettings["gph:GraphApiUrl"] }/groups?$filter=mailNickName eq '{ mailNickname }'").Result;
            var responseData = JsonConvert.DeserializeObject<GraphApiGetGroupResponse>(response);
            //log.Info($"Group exist response data {responseData}");

            if (responseData != null && responseData.GraphGroupObjects != null && responseData.GraphGroupObjects.Count > 0)
            {
                //Group exists if the count is more than one
                this.TraceWriter.Info($"Group with the name nickname { mailNickname } already exists");
                return true;
            }
            else
            {
                return false;
            }
            
        }

        public bool DoesSharePointSiteExist(string fullUrl)
        {
            //https://graph.microsoft.com/v1.0/sites/M365x386378.sharepoint.com:/sites/group01

            //Get server relative url
            string lowerFullUrl = fullUrl.ToLower();
            string relativeUrl = lowerFullUrl.Replace("https://m365x386378.sharepoint.com", String.Empty);
            this.TraceWriter.Info($"Server relative url is {relativeUrl}");

            string tenantName = ConfigurationManager.AppSettings["o365:SpoTenantUrl"].Replace("https://", string.Empty);
            string getUrl = $"{ ConfigurationManager.AppSettings["gph:GraphApiUrl"] }/sites/{ tenantName }:{ relativeUrl }";
            this.TraceWriter.Info($"Get Url is {getUrl}");

            using (HttpClient client = new HttpClient())
            {
                //Get authentication token
                string token = this.GetGraphApiDelegatedAuthenticationToken().Result;
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                var response = client.GetAsync(getUrl).Result;

                this.TraceWriter.Info($"Response status code is { response.StatusCode }");

                if (response.StatusCode == HttpStatusCode.OK)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
        
    }
}
