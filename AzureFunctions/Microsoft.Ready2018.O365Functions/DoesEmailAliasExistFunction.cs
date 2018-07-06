using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Ready2018.O365Functions.Models;
using Microsoft.Ready2018.O365Functions.Utilities;
using Newtonsoft.Json;
using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace Microsoft.Ready2018.O365Functions
{
    public static class DoesEmailAliasExistFunction
    {
        [FunctionName("DoesEmailAliasExist")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            try
            {
                var content = await req.Content.ReadAsStringAsync();
                log.Info($"Function was triggered with the following payload { content }");

                var request = JsonConvert.DeserializeObject<DoesEmailNickNameExistRequest>(content);
                var graph = new GraphApiUtility(log);

                var exists = graph.DoesUserEmailNicknameExists(request.EmailAlias).Result;
                log.Info($"{request.EmailAlias} exists is {exists}");

                return req.CreateResponse(HttpStatusCode.OK, new { emailAliasExists = exists, emailAlias = request.EmailAlias, message = exists.ToString() });
            }
            catch (Exception ex)
            {
                return req.CreateResponse(HttpStatusCode.OK, new { emailAliasExists = false, message = ex.ToDetailedString()});
            }
        }
    }
}
