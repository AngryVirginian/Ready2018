using Newtonsoft.Json;

namespace Microsoft.Ready2018.O365Functions.Models
{
    public class CreateCommunicationSiteRequest
    {
        [JsonProperty(PropertyName = "title")]
        public string Title { get; set; }

        [JsonProperty(PropertyName = "description")]
        public string Description { get; set; }

        [JsonProperty(PropertyName = "url")]
        public string Url { get; set; }

        [JsonProperty(PropertyName = "allowFileSharingForGuestUsers")]
        public bool AllowFileSharingForGuestUsers { get; set; }

        [JsonProperty(PropertyName = "siteDesign")]
        public string SiteDesign { get; set; }

        [JsonProperty(PropertyName = "owner")]
        public string Owner { get; set; }

    }
}
