using Newtonsoft.Json;

namespace Microsoft.Ready2018.O365Functions.Models
{
    /// <summary>
    /// Team settings that can be passed to Graph
    /// </summary>
    public class TeamSettings
    {
        /// <summary>
        /// Ctor
        /// </summary>
        public TeamSettings()
        {
            this.TeamMemberSettings = new TeamMemberSettings();
            this.TeamGuestSettings = new TeamGuestSettings();
            this.TeamMessagingSettings = new TeamMessagingSettings();
            this.TeamFunSettings = new TeamFunSettings();
        }

        [JsonProperty(PropertyName = "memberSettings")]
        public TeamMemberSettings TeamMemberSettings { get; set; }

        [JsonProperty(PropertyName = "guestSettings")]
        public TeamGuestSettings TeamGuestSettings { get; set; }

        [JsonProperty(PropertyName = "messagingSettings")]
        public TeamMessagingSettings TeamMessagingSettings { get; set; }

        [JsonProperty(PropertyName = "funSettings")]
        public TeamFunSettings TeamFunSettings { get; set; }

    }

    public class TeamMemberSettings
    {
        /// <summary>
        /// Ctor
        /// </summary>
        public TeamMemberSettings()
        {
            //Default.  Those not set will be false
            this.AllowCreateUpdateChannels = true;
            this.AllowCreateUpdateRemoveTabs = true;
        }

        [JsonProperty(PropertyName = "allowCreateUpdateChannels")]
        public bool AllowCreateUpdateChannels { get; set; }

        [JsonProperty(PropertyName = "allowDeleteChannels")]
        public bool AllowDeleteChannels { get; set; }

        [JsonProperty(PropertyName = "allowAddRemoveApps")]
        public bool AllowAddRemoveApps { get; set; }

        [JsonProperty(PropertyName = "allowCreateUpdateRemoveTabs")]
        public bool AllowCreateUpdateRemoveTabs { get; set; }

        [JsonProperty(PropertyName = "allowCreateUpdateRemoveConnectors")]
        public bool AllowCreateUpdateRemoveChannels { get; set; }

    }

    public class TeamGuestSettings
    {

        [JsonProperty(PropertyName = "allowCreateUpdateChannels")]
        public bool AllowCreateUpdateChannels { get; set; }

        [JsonProperty(PropertyName = "allowDeleteChannels")]
        public bool AllowDeleteChannels { get; set; }

    }

    public class TeamMessagingSettings
    {

        /// <summary>
        /// Ctor
        /// </summary>
        public TeamMessagingSettings()
        {
            //Default.  Those not set will be false
            this.AllowUserEditMessages = true;
            this.AllowUserDeleteMessages = true;
            this.AllowChannelMentions = true;
            this.AllowTeamMentions = true;
        }

        [JsonProperty(PropertyName = "allowUserEditMessages")]
        public bool AllowUserEditMessages { get; set; }

        [JsonProperty(PropertyName = "allowUserDeleteMessages")]
        public bool AllowUserDeleteMessages { get; set; }

        [JsonProperty(PropertyName = "allowOwnerDeleteMessages")]
        public bool AllowOwnerDeleteMessages { get; set; }

        [JsonProperty(PropertyName = "allowTeamMentions")]
        public bool AllowTeamMentions { get; set; }

        [JsonProperty(PropertyName = "allowChannelMentions")]
        public bool AllowChannelMentions { get; set; }
    }


    public class TeamFunSettings
    {
        /// <summary>
        /// Ctor
        /// </summary>
        public TeamFunSettings()
        {
            //Default.  Those not set will be false
            this.AllowGiphy = true;
            this.GiphyContentRating = "strict"; //default to strict
        }

        [JsonProperty(PropertyName = "allowGiphy")]
        public bool AllowGiphy { get; set; }

        [JsonProperty(PropertyName = "giphyContentRating")]
        public string GiphyContentRating { get; set; }

        [JsonProperty(PropertyName = "allowStickersAndMemes")]
        public bool AllowStickersAndMemes { get; set; }

        [JsonProperty(PropertyName = "allowCustomMemes")]
        public bool AllowCustomMemes { get; set; }
    }
}
