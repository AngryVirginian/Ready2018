using System;
using Newtonsoft.Json;

namespace Microsoft.Ready2018.O365Functions.Models
{
    public class DoesSiteExistsRequest
    {
        [JsonProperty(PropertyName = "fullUrl")]
        public string FullUrl { get; set; }
    }

}
