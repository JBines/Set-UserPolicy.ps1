# Set-UserAddressBookPolicy.ps1

This script automates applying custom Address Book Polices to a select groups of users.  

### DESCRIPTION
This has been created to run on a schedule and apply the correct Address Book Policy.  

**Set-UserPolicy.ps1 [-SourceGroups <Array[ObjectID]>] [-AddressBookPolicy <String[ABP Name]>] [-MailboxRetentionPolicy <String[ABP Name]>] [-AddressBookPolicy <String[ABP Name]>] [-EnablePublicFolderClientAccess <Switch>]**


```PowerShell
<# 
.SYNOPSIS
This script automates applying custom User Polices to a select groups of users.  

.DESCRIPTION
This has been created to run on a schedule and apply the correct Address Book Policy.  

## Set-UserPolicy.ps1 [-SourceGroups <Array[ObjectID]>] [-AddressBookPolicy <String[ABP Name]>] [-MailboxRetentionPolicy <String[ABP Name]>] [-AddressBookPolicy <String[ABP Name]>] [-EnablePublicFolderClientAccess <Switch>] 

.PARAMETER SourceGroups
The SourceGroup parameter details the ObjectId of the Azure Group which contains all the desired users that need the Address Book Policy.

.PARAMETER AddressBookPolicy
The AddressBookPolicy parameter specifies the name of the Address Book Policy which should be applied. 

.PARAMETER EnablePublicFolderClientAccess
The EnablePublicFolderClientAccess switch will enable the mailbox for Public Folder client access.

.PARAMETER DisablePublicFolderClientAccess
The DisablePublicFolderClientAccess switch will disable the mailbox for Public Folder client access.

.PARAMETER DifferentialScope
The DifferentialScope parameter defines how many objects can be added or removed from the UserGroups in a single operation of the script. The goal of this setting is throttle bulk changes to limit the impact of misconfiguration by an administrator. What value you choose here will be dictated by your userbase and your script schedule. The default value is set to 10 Objects. 

.PARAMETER AutomationPSCredential
The DifferentialScope parameter defines how many objects can be added or removed from the UserGroups in a single operation of the script. The goal of this setting is throttle bulk changes to limit the impact of misconfiguration by an administrator. What value you choose here will be dictated by your userbase and your script schedule. The default value is set to 10 Objects. 

.EXAMPLE
Set-UserPolicy.ps1 -SourceGroup '7b7c4926-c6d7-4ca8-9bbf-5965751022c2' -AddressBookPolicy 'Executive ABP'

-- SET MEMBERS FOR ROLE GROUPS --

In this example the script will apply the 'Executive ABP' to mailbox users who are members of Group '7b7c4926-c6d7-4ca8-9bbf-5965751022c2' 

.LINK

Address Book Policies, Jamba Jokes and Secret Agents - https://techcommunity.microsoft.com/t5/exchange-team-blog/address-book-policies-jamba-jokes-and-secret-agents/ba-p/595749

Assign an address book policy to users in Exchange Online - https://docs.microsoft.com/en-us/exchange/address-books/address-book-policies/assign-an-address-book-policy-to-mail-users

.NOTES
This function requires that you have already created your Dynamic Azure AD Groups.

Please note, when using Azure Automation with more than one user group the array should be set to JSON for example ['ObjectID','ObjectID']

[AUTHOR]
Joshua Bines, Consultant

Find me on:
* Web:     https://theinformationstore.com.au
* LinkedIn:  https://www.linkedin.com/in/joshua-bines-4451534
* Github:    https://github.com/jbines
  
[VERSION HISTORY / UPDATES]
0.0.1 20200121 - JBINES - Created the bare bones
0.0.2 20200127 - JBINES - Added support for Enabling and Disabling Client Access to Public Folders

[TO DO LIST / PRIORITY]

#>

```
