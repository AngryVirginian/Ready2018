using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Microsoft.Ready2018.O365Functions.Models
{
    /// <summary>
    /// Json data for Graph API Channel Create Request
    /// </summary>
    public class GraphApiChannelRequest
    {
        [JsonProperty(PropertyName = "displayName")]
        public string DisplayName { get; set; }

        [JsonProperty(PropertyName = "description")]
        public string Description { get; set; }

    }

    /// <summary>
    /// Json data for Graph API Channel object
    /// </summary>
    public class GraphApiChannel : GraphApiChannelRequest
    {
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }
    }
}
