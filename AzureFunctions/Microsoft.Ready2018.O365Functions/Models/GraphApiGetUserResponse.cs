using Newtonsoft.Json;
using System.Collections.Generic;


namespace Microsoft.Ready2018.O365Functions.Models
{
    /// <summary>
    /// Data model from the Graph API get users call
    /// </summary>
    public class GraphApiGetUserResponse
    {
        [JsonProperty(PropertyName = "@odata.context")]
        public string ODataContext { get; set; }

        [JsonProperty(PropertyName = "value")]
        public List<GraphApiUser> GraphUserObjects { get; set; }

    }
}
