using Newtonsoft.Json;

namespace Microsoft.Ready2018.O365Functions.Models
{
    public class DoesEmailNickNameExistRequest
    { 
        [JsonProperty(PropertyName = "emailAlias")]
        public string EmailAlias { get; set; }
    }

   
}
