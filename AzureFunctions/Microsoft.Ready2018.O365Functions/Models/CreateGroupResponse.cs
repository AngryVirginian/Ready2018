using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Microsoft.Ready2018.O365Functions.Models
{
    public class CreateGroupResponse
    {
        [JsonProperty(PropertyName = "groupCreated")]
        public bool GroupCreated { get; set; }

        [JsonProperty(PropertyName = "groupId")]
        public Guid GroupId { get; set; }

        [JsonProperty(PropertyName = "groupSiteUrl")]
        public string GroupSiteUrl { get; set; }

        [JsonProperty(PropertyName = "teamCreated")]
        public bool TeamCreated { get; set; }

        [JsonProperty(PropertyName = "numberOfChannelCreated")]
        public int NumberOfChannelCreated { get; set; }

        [JsonProperty(PropertyName = "createdChannels")]
        public List<CreateChannelResponse> CreatedChannels { get; set; }

        [JsonProperty(PropertyName = "plannerCreated")]
        public bool PlannerCreated { get; set; }

        [JsonProperty(PropertyName = "notebookCreated")]
        public bool NotebookCreated { get; set; }

        [JsonProperty(PropertyName = "message")]
        public string Message { get; set; }
    }

    public class CreateChannelResponse
    {

        [JsonProperty(PropertyName = "displayName")]
        public string ChannelName { get; set; }

        [JsonProperty(PropertyName = "id")]
        public string ChannelId { get; set; }
    }
}
