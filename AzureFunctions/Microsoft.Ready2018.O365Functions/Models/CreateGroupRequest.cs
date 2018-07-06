using System;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Microsoft.Ready2018.O365Functions.Models
{
    /// <summary>
    /// Json data passed into the GroupRequestWebhook Azure Function
    /// </summary>
    public class CreateGroupRequest
    {
        [JsonProperty(PropertyName = "name")]
        public string Displayname { get; set; }

        [JsonProperty(PropertyName = "description")]
        public string Description { get; set; }

        [JsonProperty(PropertyName = "emailAlias")]
        public string EmailAlias { get; set; }

        [JsonProperty(PropertyName = "emailEnabled")]
        public bool EmailEnabled { get; set; }

        [JsonProperty(PropertyName = "ownerUpn")]
        public string OwnerUpn { get; set; }

        [JsonProperty(PropertyName = "isPrivate")]
        public bool IsPrivate { get; set; }

        [JsonProperty(PropertyName = "createTeam")]
        public bool CreateTeam { get; set; }

        [JsonProperty(PropertyName = "teamSettings")]
        public TeamSettings TeamSettings { get; set; }

        [JsonProperty(PropertyName = "teamChannels")]
        public List<GraphApiChannelRequest> TeamChannels { get; set; }

        [JsonProperty(PropertyName = "createPlanner")]
        public bool CreatePlanner { get; set; }

        [JsonProperty(PropertyName = "plannerTitle")]
        public string PlannerTitle { get; set; }

        [JsonProperty(PropertyName = "createNotebook")]
        public bool CreateNotebook { get; set; }

        [JsonProperty(PropertyName = "notebookTitle")]
        public string NotebookTitle { get; set; }
    }
}
