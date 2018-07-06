using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace Microsoft.Ready2018.O365Functions.Models
{
    /// <summary>
    /// Graph Api Group data model
    /// </summary>
    /// <remarks>Graph API will return this if the call to get Group or create Group is successful</remarks>
    public class GraphApiGroup
    {
        [JsonProperty(PropertyName = "id")]
        public string ID { get; set; }

        [JsonProperty(PropertyName = "deletedDateTime")]
        public object DeletedDateTime { get; set; }

        [JsonProperty(PropertyName = "classification")]
        public object Classification { get; set; }

        [JsonProperty(PropertyName = "createdDateTime")]
        public DateTime CreatedDateTime { get; set; }

        [JsonProperty(PropertyName = "description")]
        public string Description { get; set; }

        [JsonProperty(PropertyName = "displayName")]
        public string DisplayName { get; set; }

        [JsonProperty(PropertyName = "groupTypes")]
        public List<string> GroupTypes { get; set; }

        [JsonProperty(PropertyName = "mail")]
        public string Mail { get; set; }

        [JsonProperty(PropertyName = "mailEnabled")]
        public bool MailEnabled { get; set; }

        [JsonProperty(PropertyName = "mailNickname")]
        public string MailNickname { get; set; }

        [JsonProperty(PropertyName = "onPremisesLastSyncDateTime")]
        public object OnPremisesLastSyncDateTime { get; set; }

        [JsonProperty(PropertyName = "onPremisesProvisioningErrors")]
        public List<object> OnPremisesProvisioningErrors { get; set; }

        [JsonProperty(PropertyName = "onPremisesSecurityIdentifier")]
        public object OnPremisesSecurityIdentifier { get; set; }

        [JsonProperty(PropertyName = "onPremisesSyncEnabled")]
        public object OnPremisesSyncEnabled { get; set; }

        [JsonProperty(PropertyName = "preferredDataLocation")]
        public object PreferredDataLocation { get; set; }

        [JsonProperty(PropertyName = "proxyAddresses")]
        public List<string> ProxyAddresses { get; set; }

        [JsonProperty(PropertyName = "renewedDateTime")]
        public DateTime RenewedDateTime { get; set; }

        [JsonProperty(PropertyName = "securityEnabled")]
        public bool SecurityEnabled { get; set; }

        [JsonProperty(PropertyName = "visibility")]
        public string Visibility { get; set; }
    }
}
