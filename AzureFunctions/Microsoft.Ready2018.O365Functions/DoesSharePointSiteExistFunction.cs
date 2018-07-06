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
using System;
namespace Microsoft.Ready2018.O365Functions
{
    public static class DoesSharePointSiteExistFunction
    {
        [FunctionName("DoesSharePointSiteExist")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log,  ExecutionContext executionContext)
        {
            try
            { 
                var content = await req.Content.ReadAsStringAsync();
                log.Info($"Function was triggered with the following payload { content }");

                var request = JsonConvert.DeserializeObject<DoesSiteExistsRequest>(content);
                bool exists = new GraphApiUtility(log).DoesSharePointSiteExist(request.FullUrl);

                log.Info($"Site { request.FullUrl} exists is { exists }");

                return req.CreateResponse(HttpStatusCode.OK, new { siteExists = exists, message = exists.ToString() });

            }
            catch (Exception ex)
            {
                log.Info(ex.ToDetailedString());
                return (req.CreateErrorResponse(HttpStatusCode.InternalServerError, ex));
            }

        }
    }
}
