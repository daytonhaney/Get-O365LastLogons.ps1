#Connect to Microsoft Graph
Connect-MgGraph -Scopes "AuditLog.Read.All","User.Read.All"
 
#Set the Graph Profile
Select-MgProfile beta
 
#Properties to Retrieve
$Properties = @(
    'Id','DisplayName','Mail','UserPrincipalName','UserType', 'AccountEnabled', 'SignInActivity'   
)
 
#Get All users along with the properties
$AllUsers = Get-MgUser -All -Property $Properties #| Select-Object $Properties
 
$SigninLogs = @()
ForEach ($User in $AllUsers)
{
    $SigninLogs += [PSCustomObject][ordered]@{
            LoginName       = $User.UserPrincipalName
            Email           = $User.Mail
            DisplayName     = $User.DisplayName
            UserType        = $User.UserType
            AccountEnabled  = $User.AccountEnabled
            LastSignIn      = $User.SignInActivity.LastSignInDateTime
    }
}
 
$SigninLogs
 
#Export Data to CSV
$SigninLogs | Export-Csv -Path "C:\Temp\SigninLogs.csv" -NoTypeInformation
