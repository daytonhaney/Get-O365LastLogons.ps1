<meta name="google-site-verification" content="9GWpf3mfdGfzrsqkEuPXMckE_O7e35wrXJ5P9YDcSIQ" />  

# Get-O365LastLogons.ps1  
PowerShell Script to get Office 365 last logons for Users and Guests.  

Author       : Cary GARVIN  
Contact      : cary(at)garvin.tech  
LinkedIn     : https://www.linkedin.com/in/cary-garvin  
GitHub       : https://github.com/carygarvin/  


Script Name  : Get-O365LastLogons.ps1  
Version      : 1.0  
Release date : 05/05/2020 (CET)  

History      : The present script has been developed for Organizations to have an audit view on last actions by users and guests in their Office 365 tenant.  

Purpose      : The present Script generates a list of Office365 last logons along with basic information such as Workload (O365 product), Client IP Address and so on.  
               The Script will output 2 CSV files, one with Last Logons for Office365 Users (differentiated on 'UserType' property) and another one for Office 365 Guests.  

# Script usage
This script is to be launched within "Exchange Online PowerShell" in order to invoke the cmdlet '`Search-UnifiedAuditLog`' around which the present Script is built.  
Running it from a PS-Session is not advised in case the User Management Admin account used to run it is subject to MFA.  
Supply your O365 User Management Admin credentials to the MsolService when prompted.  
Auditing for the O365 tenant should be enabled otherwise Unified Audit Logs could potentially not contain any worthwhile information...  
