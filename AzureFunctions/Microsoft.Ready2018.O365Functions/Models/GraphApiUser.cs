using Newtonsoft.Json;
using System.Collections.Generic;

namespace Microsoft.Ready2018.O365Functions.Models
{
    /// <summary>
    /// User object as returned from Graph api
    /// </summary>
    public class GraphApiUser
    {
        [JsonProperty(PropertyName = "id")]
        public string ID { get; set; }

        [JsonProperty(PropertyName = "businessPhones")]
        public List<string> BusinessPhones { get; set; }

        [JsonProperty(PropertyName = "displayName")]
        public string DisplayName { get; set; }

        [JsonProperty(PropertyName = "givenName")]
        public string GivenName { get; set; }

        [JsonProperty(PropertyName = "jobTitle")]
        public object JobTitle { get; set; }

        [JsonProperty(PropertyName = "mail")]
        public string Mail { get; set; }

        [JsonProperty(PropertyName = "mobilePhone")]
        public string MobilePhone { get; set; }

        [JsonProperty(PropertyName = "officeLocation")]
        public object OfficeLocation { get; set; }

        [JsonProperty(PropertyName = "preferredLanguage")]
        public string PreferredLanguage { get; set; }

        [JsonProperty(PropertyName = "surname")]
        public string Surname { get; set; }

        [JsonProperty(PropertyName = "userPrincipalName")]
        public string UserPrincipalName { get; set; }
    }
}
