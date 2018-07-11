using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Ready2018.O365Functions.Models;
using Microsoft.Ready2018.O365Functions.Utilities;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Sites;
using OfficeDevPnP.Core.Entities;
using System.Configuration;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;

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

                //Connect to the Tenant Root Url
                using (ClientContext clientContext = new SharePointUtility(log).GetClientContext(ConfigurationManager.AppSettings["o365:SpoTenantUrl"], executionContext))
                {

                    var tenant = new Tenant(clientContext);
                    log.Info($"Site { request.Url} exists is {tenant.SiteExists(request.Url)}");

                    log.Info($"Preparing site creation info to { request.Url }");

                    // Create new "modern" communication site using PnP
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

                    //Somehow the codes above does not add request.Owner as site collection admin
                    using (ClientContext newSiteClientContext = new SharePointUtility(log).GetClientContext(communicationContext.Web.Url, executionContext))
                    {
                        var siteCollectionAdmins = new List<UserEntity>();
                        siteCollectionAdmins.Add(new UserEntity() { LoginName = request.Owner, Email = request.Owner });
                        newSiteClientContext.Web.AddAdministrators(siteCollectionAdmins, true);
                        log.Info($"{ request.Owner} was added as site collection admin");
                    }
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
