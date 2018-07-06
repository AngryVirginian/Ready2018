Import-Module AzureADPreview

Connect-AzureAD
$securityGroupName = "AllowedtoCreateUnifiedGroup"
Get-AzureADGroup -SearchString $securityGroupName
$Template = Get-AzureADDirectorySettingTemplate | where {$_.DisplayName -eq 'Group.Unified'}
$Setting = $Template.CreateDirectorySetting()
New-AzureADDirectorySetting -DirectorySetting $Setting
$Setting = Get-AzureADDirectorySetting -Id (Get-AzureADDirectorySetting | where -Property DisplayName -Value "Group.Unified" -EQ).id
$Setting["EnableGroupCreation"] = $False
$Setting["GroupCreationAllowedGroupId"] = (Get-AzureADGroup -SearchString $securityGroupName).objectid
Set-AzureADDirectorySetting -Id (Get-AzureADDirectorySetting | where -Property DisplayName -Value "Group.Unified" -EQ).id -DirectorySetting $Setting

#To verify the security group CAN create groups, and everyone else can't:
(Get-AzureADDirectorySetting).Values
