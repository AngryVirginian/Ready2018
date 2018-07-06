namespace Microsoft.Ready2018.O365Functions.Models
{
    /// <summary>
    /// The authentication mode that the function will use to authenticate with SPO
    /// </summary>
    public enum SpoAuthenticationMethod
    {
        AzureAppIdentity,
        UserNamePassword,
        SharePointAppIdentity,
    }
}
