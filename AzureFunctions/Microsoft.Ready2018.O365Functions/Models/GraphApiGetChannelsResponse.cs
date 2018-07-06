using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace Microsoft.Ready2018.O365Functions.Models
{
    public class GraphApiGetChannelsResponse
    {
        [JsonProperty(PropertyName = "@odata.context")]
        public string Context { get; set; }

        [JsonProperty(PropertyName = "value")]
        public List<GraphApiChannel> Channels { get; set; }

    }
}
