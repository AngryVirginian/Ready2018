using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Sites;
using Microsoft.SharePoint.Client;
using Microsoft.Ready2018.O365Functions.Utilities;
using Newtonsoft.Json;
using Microsoft.Ready2018.O365Functions.Models;
using System.Configuration;
using Microsoft.Online.SharePoint.TenantAdministration;

namespace Microsoft.Ready2018.O365Functions
{
    public static class CreateCommunicationSiteFunction
    {
        [FunctionName("CreateCommunicationSite")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log, ExecutionContext executionContext)
        {
            try
            {
                var content = await req.Content.ReadAsStringAsync();
                log.Info($"Function was triggered with the following payload { content }");

                var request = JsonConvert.DeserializeObject<CreateCommunicationSiteRequest>(content);

                //var token = new GraphApiUtility(log).GetGraphApiDelegatedAuthenticationToken(ConfigurationManager.AppSettings["o365:SpoTenantUrl"]).Result;
                //using (ClientContext clientContext = AppForSharePointOnlineWebToolkit.TokenHelper.GetClientContextWithAccessToken(ConfigurationManager.AppSettings["o365:SpoTenantUrl"], token))

                using (ClientContext clientContext = new SharePointUtility(log).GetClientContext(ConfigurationManager.AppSettings["o365:SpoTenantUrl"], executionContext))
                {

                    var tenant = new Tenant(clientContext);
                    log.Info($"Site { request.Url} exists is {tenant.SiteExists(request.Url)}");

                    log.Info($"Preparing site creation info to { request.Url }");
                    // Create new "modern" communication site at the url https://[tenant].sharepoint.com/sites/mymoderncommunicationsite
                    var communicationContext = await clientContext.CreateSiteAsync(new CommunicationSiteCollectionCreationInformation
                    {
                        Url = request.Url, // Mandatory
                        Title = request.Title, // Mandatory
                        Description = request.Description, // Mandatory
                        Owner = request.Owner, // Optional
                        AllowFileSharingForGuestUsers = request.AllowFileSharingForGuestUsers, // Optional
                        SiteDesign = CommunicationSiteDesign.Topic, // Mandatory
                        Lcid = 1033, // Mandatory 
                        Classification = "classification", // Optional

                    });
                    communicationContext.Load(communicationContext.Web, w => w.Url);
                    communicationContext.ExecuteQueryRetry();
                    log.Info($"New communication site created at {communicationContext.Web.Url}");
                    //Console.WriteLine(communicationContext.Web.Url);
                }

                return req.CreateResponse(HttpStatusCode.OK, new { siteCreated = true, siteUrl = request.Url });
            }
            catch (System.Exception e)
            {
                log.Info(e.ToString());
                return req.CreateResponse(HttpStatusCode.InternalServerError, new { siteCreated = false, message = e.ToString() });
                
            }
        }
    }
}
