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
0.0.1 20200121 - JBINES - Created the bare bones.
0.0.2 20200127 - JBINES - Added support for Enabling and Disabling Client Access to Public Folders.
0.0.3 20200127 - JBINES - Changed Switch to boolean for Azure Automation Best Practices.

[TO DO LIST / PRIORITY]

#>

Param 
(
    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [array]$SourceGroups,
    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [string]$AddressBookPolicy,
    [Parameter(Mandatory = $False)]
    [ValidateNotNullOrEmpty()]
    [Int]$DifferentialScope = 10,
    [Parameter(Mandatory = $False)]
    [boolean]$EnablePublicFolderClientAccess=$False,
    [Parameter(Mandatory = $False)]
    [boolean]$DisablePublicFolderClientAccess=$False,
    [Parameter(Mandatory = $False)]
    [ValidateNotNullOrEmpty()]
    [String]$AutomationPSCredential
)

    #Set VAR
    $counter = 0

# Success Strings
    $sString0 = "CMDlet:Set-Mailbox"

    # Info Strings
    $iString0 = "Set Address Book Policy"

# Warn Strings
    $wString0 = "CMDlet:Measure-Object;No Members"

# Error Strings

    $eString1 = "Hey! You hit the -DifferentialScope limit. Let's break out of this loop"
    $eString2 = "Hey! Looks like we are having issues finding your ABP or found more than on policy"

    function Write-Log([string[]]$Message, [string]$LogFile = $Script:LogFile, [switch]$ConsoleOutput, [ValidateSet("SUCCESS", "INFO", "WARN", "ERROR", "DEBUG")][string]$LogLevel)
    {
           $Message = $Message + $Input
           If (!$LogLevel) { $LogLevel = "INFO" }
           switch ($LogLevel)
           {
                  SUCCESS { $Color = "Green" }
                  INFO { $Color = "White" }
                  WARN { $Color = "Yellow" }
                  ERROR { $Color = "Red" }
                  DEBUG { $Color = "Gray" }
           }
           if ($Message -ne $null -and $Message.Length -gt 0)
           {
                  $TimeStamp = [System.DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
                  if ($LogFile -ne $null -and $LogFile -ne [System.String]::Empty)
                  {
                         Out-File -Append -FilePath $LogFile -InputObject "[$TimeStamp] [$LogLevel] $Message"
                  }
                  if ($ConsoleOutput -eq $true)
                  {
                         Write-Host "[$TimeStamp] [$LogLevel] :: $Message" -ForegroundColor $Color

                    if($AutomationPSCredential)
                    {
                         Write-Output "[$TimeStamp] [$LogLevel] :: $Message"
                    } 
                  }
           }
    }

    #Validate Input Values From Parameter 

    Try{

        if ($AutomationPSCredential) {
            
            $Credential = Get-AutomationPSCredential -Name $AutomationPSCredential

            Connect-AzureAD -Credential $Credential
            
            #$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
            #Import-PSSession $Session -DisableNameChecking -Name ExSession -AllowClobber:$true | Out-Null

            $ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection -Name $ConnectionName 
            Import-Module (Import-PSSession -Session $ExchangeOnlineSession -AllowClobber -DisableNameChecking) -Global

            }
                            
        $objSourceGroupMembers = @($SourceGroups | ForEach-Object {Get-AzureADGroupMember -ObjectId $_})

        #Return Only Unique values remove any duplicates
        $SourceGroupMembers = $objSourceGroupMembers | Select-Object -Unique
        
        #Confirm Address Book Policy and obtain the DN. 
        $objAddressBookPolicy = Get-AddressBookPolicy $AddressBookPolicy
    }
    
    Catch{
    
        $ErrorMessage = $_.Exception.Message
        Write-Error $ErrorMessage

            If($?){Write-Log -Message $ErrorMessage -LogLevel Error -ConsoleOutput}

        Break

    }
    
    if (($objAddressBookPolicy|Measure-Object).count -eq 1) {
        
        Write-Log -Message "$iString0 - $($objAddressBookPolicy.Name)" -LogLevel INFO -ConsoleOutput
        
        #Foreach Source group members
        foreach ($user in $SourceGroupMembers) {
            If($user.MailNickName){
                $userMailbox = $null
                $userObjectId = $null
                $objAddressBookPolicyDistinguishedName = $null
                $userObjectId = $user.ObjectId
                $objAddressBookPolicyDistinguishedName = $objAddressBookPolicy.DistinguishedName
                $userMailbox = Get-Mailbox -Filter "RecipientTypeDetails -eq 'UserMailbox' -and ExternalDirectoryObjectId -eq '$userObjectId' -and AddressBookPolicy -ne '$objAddressBookPolicyDistinguishedName'"

                if ($userMailbox) {
                    
                    if ($counter -lt $DifferentialScope) {

                        Set-Mailbox $user.ObjectId -AddressBookPolicy $objAddressBookPolicy.Name
                        if($?){Write-Log -Message "$sString0;ABP:$($objAddressBookPolicy.Name);UserObjectId:$($user.ObjectId);UserUPN:$($user.UserPrincipalName)" -LogLevel SUCCESS -ConsoleOutput}

                        if($EnablePublicFolderClientAccess){

                            Set-CASMailbox $user.ObjectId -PublicFolderClientAccess:$True -Confirm:$false
                            if($?){Write-Log -Message "$sString0;PFClientAcceess:TRUE;UserObjectId:$($user.ObjectId);UserUPN:$($user.UserPrincipalName)" -LogLevel SUCCESS -ConsoleOutput}
                        }

                        if($DisablePublicFolderClientAccess){

                            Set-CASMailbox $user.ObjectId -PublicFolderClientAccess:$False -Confirm:$false
                            if($?){Write-Log -Message "$sString0;PFClientAcceess:False;UserObjectId:$($user.ObjectId);UserUPN:$($user.UserPrincipalName)" -LogLevel SUCCESS -ConsoleOutput}
                        }

                        $counter++
                    }
                    
                    else {
                        
                        #Exceeded couter limit
                        Write-log -Message $eString1 -ConsoleOutput -LogLevel ERROR
                        Break
                        
                    }
                }
            }
        }
    }
    else {
        Write-Log -Message $eString2 -LogLevel Error -ConsoleOutput
    }

if ($AutomationPSCredential) {
    
    #Invoke-Command -Session $ExchangeOnlineSession -ScriptBlock {Remove-PSSession -Session $ExchangeOnlineSession}

    Disconnect-AzureAD
}
