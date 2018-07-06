using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using Microsoft.Ready2018.O365Functions.Models;
using Microsoft.Ready2018.O365Functions.Utilities;
using System;
using OfficeDevPnP.Core.Framework.Graph;
using System.Text;
using System.Net.Http.Headers;
using System.Configuration;
using System.Collections.Generic;

namespace Microsoft.Ready2018.O365Functions
{
    public static class CreateGroupFunction
    {
        [FunctionName("CreateGroup")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            //Set up the results to return to caller
            CreateGroupResponse result = new CreateGroupResponse() { CreatedChannels = new List<CreateChannelResponse>() };

            try
            {
                var content = await req.Content.ReadAsStringAsync();
                log.Info($"Function was triggered with the following payload { content }");

                try
                {
                    var request = JsonConvert.DeserializeObject<CreateGroupRequest>(content);
                    log.Info("Successfully deserialized Json data");
                    
                    //Get Azure OAuth access token
                    string accessToken = await (new GraphApiUtility(log)).GetGraphApiAuthenticationToken();

                    List<string> owners = new List<string>();
                    owners.Add(request.OwnerUpn);
                    // Graph API create team only supports delegated access permission.  Must add service principal to owner of the Group
                    string teamServicePrincipal = KeyVaultUtility.GetSecret(ConfigurationManager.AppSettings["kv:ServicePrincipalNameSecretName"], log).Result;
                    if (teamServicePrincipal.ToLower() != request.OwnerUpn.ToLower())
                    {
                        owners.Add(teamServicePrincipal);
                    }

                    //Use Pnp to create a new Group
                    var newGroup = UnifiedGroupsUtility.CreateUnifiedGroup(displayName: request.Displayname, 
                        description: request.Description, 
                        mailNickname: request.EmailAlias, 
                        accessToken: accessToken,
                        owners: owners.ToArray(),
                        members: owners.ToArray(),
                        isPrivate: request.IsPrivate,
                        retryCount: 10,
                        delay: 500);

                    //Set up the results to return to caller
                    result.GroupCreated = true;
                    result.GroupId = new Guid(newGroup.GroupId);
                    result.GroupSiteUrl = newGroup.SiteUrl;

                    if (request.CreateTeam)
                    {
                        //Add apps such as Team (& Channel), Planner, and OneNote to Group
                        await AddAppsToTeam(request, result.GroupId, result, log);
                    }

                    return new HttpResponseMessage()
                    {
                        StatusCode = HttpStatusCode.OK,
                        Content = new StringContent(JsonConvert.SerializeObject(result), System.Text.Encoding.UTF8, "application/json")
                    };

                }
                catch (JsonReaderException jsonError)
                {
                    log.Error($"Json parsing error in incoming request: {jsonError.Message}");

                    result.GroupCreated = false;
                    result.Message = $"Invalid Json request: { jsonError.Message }";

                    return new HttpResponseMessage()
                    {
                        StatusCode = HttpStatusCode.InternalServerError,
                        Content = new StringContent(JsonConvert.SerializeObject(result), System.Text.Encoding.UTF8, "application/json")
                    };
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.ToDetailedString());
                result.GroupCreated = false;
                result.Message = $"Error: { ex.ToDetailedString() }";
                return new HttpResponseMessage()
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    Content = new StringContent(JsonConvert.SerializeObject(result), System.Text.Encoding.UTF8, "application/json")
                };
            }
        }

        /// <summary>
        /// Add apps such as Team (& Channel), Planner, and OneNote to Group
        /// </summary>
        private async static Task AddAppsToTeam(CreateGroupRequest request, Guid groupId, CreateGroupResponse result, TraceWriter log)
        {
            //If options to create additional app
            if (request.CreateTeam || request.CreateNotebook || request.CreatePlanner)
            {
                int pauseTime = int.Parse(ConfigurationManager.AppSettings["gph:PauseTimeAfterGroupCreationInMilliseonds"]);
                log.Info($"Pausing for { pauseTime } milliseconds after Group Creation");
                System.Threading.Thread.Sleep(pauseTime);

                using (HttpClient delegatedAccessClient = new HttpClient())
                {
                    //Get deletgate access token
                    string delegatedToken = await new GraphApiUtility(log).GetGraphApiDelegatedAuthenticationToken();
                    delegatedAccessClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", delegatedToken);

                    if (request.CreateTeam)
                    {
                        //Set the team settings to default if none was passed in
                        if (request.TeamSettings == null ) { request.TeamSettings = new TeamSettings(); }
                        //Create a team
                        result.TeamCreated = await AddTeamToUnifiedGroup(delegatedAccessClient, result.GroupId, request.TeamSettings, log);

                        //Create channel if the request has channel request and Team was created
                        if (result.TeamCreated && request.TeamChannels != null && request.TeamChannels.Count > 0)
                        {
                            for (int i = 0; i < request.TeamChannels.Count; i++)
                            {
                                if (!string.IsNullOrEmpty(request.TeamChannels[i].DisplayName))
                                {
                                    var channel = await AddChannelsToUnifiedGroup(delegatedAccessClient, result.GroupId, request.TeamChannels[i], log);
                                    if (channel != null)
                                    {
                                        result.NumberOfChannelCreated++;
                                    }
                                }
                            }

                            //Add default chat thread
                            var channels = await GetChannelsFromGroup(delegatedAccessClient, result.GroupId, log);
                            if (channels != null && channels.Count > 0)
                            {
                                for (int i = 0; i < channels.Count; i++)
                                {
                                    await AddDefaultChatThreadToTeamChannel(delegatedAccessClient, result.GroupId, channels[i], log);
                                }
                            }
                        }
                    }

                    if (request.CreatePlanner)
                    {
                        result.PlannerCreated = await AddPlannerToUnifiedGroup(delegatedAccessClient, result.GroupId, request.PlannerTitle, log);
                    }

                    if (request.CreateNotebook)
                    {
                        result.NotebookCreated = await AddNotebookToUnifiedGroup(delegatedAccessClient, result.GroupId, request.NotebookTitle, log);
                    }
                }
            }
        }




        /// <summary>
        /// Add Team to an existing unified group.  Not yet completed
        /// </summary>
        /// <returns>true if the operation is successful</returns>
        /// <remarks>Adding Team through Graph is currently in Beta. APIs under the /beta version in Microsoft Graph are in preview and are subject to change. Use of these APIs in production applications is not supported.
        private static async Task<bool> AddTeamToUnifiedGroup(HttpClient delegatedAccessClient, Guid groupId, TeamSettings teamSettings, TraceWriter log)
        {
            //TODO: only support Work Delegated permission in the beta reference.  Team is not in GCC yet.  The current Graph URL is
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/team_put_teams
            //https://microsoftteams.uservoice.com/forums/555103-public/suggestions/16972258-developer-api-to-read-create-teams-and-channels-in?page=2&per_page=20

            log.Info($"Adding Team to Group { groupId }");

            string putUrl = $"https://graph.microsoft.com/beta/groups/{groupId}/team";
            var payload = JsonConvert.SerializeObject(teamSettings);
            HttpContent content = new StringContent(payload, Encoding.UTF8, "application/json");
            //Put command
            var response = await delegatedAccessClient.PutAsync(putUrl, content); 
            if (response.IsSuccessStatusCode)
            {
                log.Info($"Added Team to group {groupId}");
                return true;
            }
            else
            {
                log.Info($"FAILED to add Team to group {groupId} with status code { response.StatusCode } ");
                log.Info(response.Content.ReadAsStringAsync().Result);
                return false;
            }
        }
        

        private static async Task<bool> AddNotebookToUnifiedGroup(HttpClient delegatedAccessClient, Guid groupId, string notebookDisplayName, TraceWriter log)
        {
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/onenote_post_notebooks

            //TODO: only support Work Delegated permission in the beta reference.  Team is not in GCC yet.  The current Graph URL is

            log.Info($"Adding Planner to Group { groupId }");

            string postUrl = $"https://graph.microsoft.com/beta/groups/{ groupId.ToString() }/onenote/notebooks";
            //log.Info($"Put url is {putUrl}");
            var postPayLoad = new
            {
                displayName = notebookDisplayName
            };

            var payloadString = JsonConvert.SerializeObject(postPayLoad);
            HttpContent content = new StringContent(payloadString, Encoding.UTF8, "application/json");
            //Put command
            var response = await delegatedAccessClient.PostAsync(postUrl, content);
            if (response.IsSuccessStatusCode)
            {
                log.Info($"Added Notebook { notebookDisplayName } to group {groupId}");
                return true;
            }
            else
            {
                log.Info($"FAILED to add Notebook to group {groupId} with status code { response.StatusCode } ");
                log.Info(response.Content.ReadAsStringAsync().Result);
                return false;
            }

        }

        private static async Task<bool> AddPlannerToUnifiedGroup(HttpClient delegatedAccessClient, Guid groupId, string plannerTitle, TraceWriter log)
        {
            //TODO: only support Work Delegated permission in the beta reference.  Team is not in GCC yet.  The current Graph URL is
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/planner_post_plans
            //POST https://graph.microsoft.com/beta/planner/plans

            log.Info($"Adding Planner to Group { groupId }");

            string postUrl = $"https://graph.microsoft.com/beta/planner/plans";
            //log.Info($"Put url is {putUrl}");
            var postPayLoad = new
            {
                owner = groupId.ToString(),
                title = plannerTitle
            };

            var payloadString = JsonConvert.SerializeObject(postPayLoad);
            HttpContent content = new StringContent(payloadString, Encoding.UTF8, "application/json");
            //Put command
            var response = await delegatedAccessClient.PostAsync(postUrl, content);
            if (response.IsSuccessStatusCode)
            {
                log.Info($"Added Planner to group {groupId}");
                return true;
            }
            else
            {
                log.Info($"FAILED to add Planner to group {groupId} with status code { response.StatusCode } ");
                log.Info(response.Content.ReadAsStringAsync().Result);
                return false;
            }
        }

        /// <summary>
        /// Add Channel to Unified Group
        /// </summary>
        /// <remarks>Adding Channel through Graph is currently in Beta. APIs under the /beta version in Microsoft Graph are in preview and are subject to change. Use of these APIs in production applications is not supported.
        /// <returns>Channel object.  Null if operation failed</returns>
        private static async Task<GraphApiChannel> AddChannelsToUnifiedGroup(HttpClient delegatedAccessClient, Guid groupId, GraphApiChannelRequest channel, TraceWriter log)
        {
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_post_channels
            //POST /groups/{id}/team/channels

            string postUrl = $"https://graph.microsoft.com/beta/groups/{ groupId.ToString() }/team/channels";
            var payload = JsonConvert.SerializeObject(channel);
            HttpContent content = new StringContent(payload, Encoding.UTF8, "application/json");

            var response = await delegatedAccessClient.PostAsync(postUrl, content);

            if (response.IsSuccessStatusCode)
            {
                log.Info($"Added Channel { channel.DisplayName } to group {groupId}");
                //var rawResult = response.Content.ReadAsStringAsync().Result;
                //log.Info($"Raw result is { rawResult}");
                return JsonConvert.DeserializeObject<GraphApiChannel>(response.Content.ReadAsStringAsync().Result);

            }
            else
            {
                log.Info($"FAILED to add channel { channel.DisplayName} to group {groupId} with status code { response.StatusCode } ");
                log.Info(response.Content.ReadAsStringAsync().Result);
                return null;
            }
            
        }

        /// <summary>
        /// Add default chat thread to Channel in a Team
        /// </summary>
        /// <returns></returns>
        /// <remarks>Adding Chat Thread through Graph is currently in Beta. APIs under the /beta version in Microsoft Graph are in preview and are subject to change. Use of these APIs in production applications is not supported.
        /// </remarks>
        private static async Task<bool> AddDefaultChatThreadToTeamChannel(HttpClient delegatedAccessClient, Guid groupId, GraphApiChannel channel, TraceWriter log)
        {
            //Add default chat thread to channel
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/channel_post_chatthreads
            string postUrl = $"https://graph.microsoft.com/beta/groups/{ groupId }/team/channels/{ channel.Id }/chatthreads";

            log.Info($"Chat create post URL is {postUrl}");
            var chatThread = new
            {
                rootMessage = new
                {
                    body = new
                    {
                        contentType = 1,
                        content = $"<H1>{channel.DisplayName}.</H1> Welcome to Ready 2018 Las Vegas!"
                    }
                }
            };
            var payload = JsonConvert.SerializeObject(chatThread);
            log.Info($"Payload is { payload }");
            var content = new StringContent(payload, Encoding.UTF8, "application/json");
            var response = await delegatedAccessClient.PostAsync(postUrl, content);

            if (response.IsSuccessStatusCode)
            {
                log.Info($"Added default chat thread to channel { channel.DisplayName }");
                return true;
            }
            else
            {
                log.Info($"FAILED to add default chatthred to {channel.DisplayName} with status code { response.StatusCode } ");
                return false;
            }
        }

        private static async Task<List<GraphApiChannel>> GetChannelsFromGroup(HttpClient delegatedAccessClient, Guid groupId, TraceWriter log)
        {
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_list_channels

            List<GraphApiChannel> channels = new List<GraphApiChannel>();

            string getUrl = $"https://graph.microsoft.com/beta/groups/{groupId}/team/channels";
            var response = await delegatedAccessClient.GetAsync(getUrl);

            var result = JsonConvert.DeserializeObject<GraphApiGetChannelsResponse>(response.Content.ReadAsStringAsync().Result);

            return result.Channels;
            
        }
    }
}
