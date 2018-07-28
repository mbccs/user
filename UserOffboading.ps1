. "C:\Scripts\UserFunctions.ps1"


###########################
# Reset Variables

$adminCredential = $null
$isDomain = $null
$isCloudOnly = $null
$isTargetUser = $null
$isUPN = $null
$isOffice365 = $null
$isAADSync = $null
$isAADConnectServer = $null
$isEXO = $null
$isExchange = $null
$isSPO = $null
$isODB = $null
$mailboxIntent = $null
$intentResponse = $null
$isRedirect = $null
$redirectEmail = $null
$isForward = $null
$isForwardEmail = $null
$ForwardEmail = $null
$isLeaveCopy = $null
$isDelegated = $null
$isDelegateUser = $null
$delegateUser = $null
$isDelegateUPN = $null
$delegateUPN =$null
$isRestricted = $null
$isHidden = $null
$isPath = $null
$isAutoResponse = $null
$ReponderType = $null
$isAutoResponseMessage = $null
$AutoResponseMessage = $null


###########################

if (!$isDomain) { Write-Output "`nIs the offboarding target user on Active Directory?"; $isDomain = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isDomain) { $isDomain = Read-Host "Y/N" } }
if ($isDomain -match "N") { if (!$isCloudOnly) { Write-Output "`nIs the offboarding target user a cloud-only (Azure AD) user?"; $isCloudOnly = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isCloudOnly) { $isCloudOnly = Read-Host "Y/N" } } }

if ($isDomain -match "Y") {
    try { Import-Module ActiveDirectory }
    catch { Write-Output "User Offboarding Script must be executed from a server with ActiveDirectory module available.  Please re-run from a Domain Controller or install RSAT Tools. "; exit; }
    
    while ($isTargetUser -notmatch "Y") {
        $isTargetUser = $Null
        Write-Output "`nSelect the target from the list.."; $targetUser = Get-ADUser -Filter * | select Enabled,Name,SamAccountName,UserPrincipalName | Out-GridView -PassThru
        if (!$isTargetUser) { Write-Output "Please confirm target user: $($targetUser.Name)"; $isTargetUser = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isTargetUser) { $isTargetUser = Read-Host "Y/N" } }
    }
    $userName = $targetUser.SamAccountName
    $upn = $targetUser.UserPrincipalName
}
elseif ($isCloudOnly -match "Y") {
    if (!$adminCredential) { 
        $adminCredential = ConnectMsol 
        if (!(Get-MsolCompanyInformation -ErrorAction SilentlyContinue)){
            Write-Output "Unable to connect to Azure AD using provided credentials. Exiting."; exit
        }
    }
    while ($isTargetUser -notmatch "Y") {
        $isTargetUser = $Null
        Write-Output "`nSelect the target from the list.."; $targetUser = Get-MsolUser | Out-GridView -PassThru
        if (!$isTargetUser) { Write-Output "Please confirm target user: $($targetUser.DisplayName)"; $isTargetUser = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isTargetUser) { $isTargetUser = Read-Host "Y/N" } }
    }
    $upn = $targetUser.UserPrincipalName
    $userName = $upn.split("@")[0]
    $isOffice365 = "Y"
    $isAADSync = "N"
    $isEXO = "Y"
    $isExchange = "N"
}
else {
    Write-Output "`nOffboarding Users who are not on Active Directory OR Azure Active Diretory is not supported."; exit
}

$transcriptpath = ".\User_Offboarding_" + $userName + "_" +(Get-Date).ToString('yyyy-MM-dd') + ".txt"
Start-Transcript -Path $transcriptpath

if (!$isOffice365) { Write-Output "`nDoes this user use Office 365?"; $isOffice365 = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isOffice365) { $isOffice365 = Read-Host "Y/N" } }

if ($isOffice365 -match "Y") {
    if (!$adminCredential) { 
        $adminCredential = ConnectMsol 
        if (!(Get-MsolCompanyInformation -ErrorAction SilentlyContinue)){
            Write-Output "Unable to connect to Azure AD using provided credentials. Exiting."; exit
        }
    }
    if (!$isAADSync) { Write-Output "`nDoes this environment use Azure AD Connect?"; $isAadsync = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isAADSync) { $isAADSync = Read-Host "Y/N" } }
    if ($isAADSync -match "Y") {
        while (!$isAADConnectServer) {
            Write-Output "`nWhat is the host/fqdn of the AD Connect Server?"
            $AADConnectServer = Read-Host -Prompt "Hostname"
            try { $isAADConnectServer = Invoke-Command -ComputerName $AADConnectServer -ScriptBlock { Get-ADSyncConnector } }
            catch { Write-Output "Unable to confirm AD Sync is configured on: $AADConnectServer" }
        }
    }
    if (!$isEXO) { Write-Output "`nDoes this user use Exchange Online? "; $isEXO = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isEXO) { $isEXO = Read-Host "Y/N" } }
    if (!$isSPO) { Write-Output "`nDoes this environment use SharePoint Online? "; $isSPO = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isSPO) { $isSPO = Read-Host "Y/N" } }
    if (!$isODB) { Write-Output "`nDoes this user use OneDrive for Business?"; $isODB = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isODB) { $isODB = Read-Host "Y/N" } }
}

if ($isEXO -match "Y") {
    if ($adminCredential){
        try {
            Write-Output "`nConnecting to Exchange Online Powershell Service..."
            $ExoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $adminCredential -Authentication Basic -AllowRedirection
            Import-PSSession $ExoSession
        }
        catch { Write-Output "! Unable to open Exchange Online session on with provided credentials"; exit; }
        #try {
        #    Write-Output "Connecting to Security and Compliance Powershell Service"
        #    $EopSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $adminCredential -Authentication Basic -AllowRedirection
        #    Import-PSSession $EopSession -AllowClobber
        #}
        #catch { Write-Output "! Unable to open Exchange Compliance session on with provided credentials" }

    }
}
else {
    if (!$isExchange) { Write-Output "`nDoes this environment use Exchange on-premise?"; $isExchange = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isExchange) { $isExchange = Read-Host "Y/N" } }
    if ($isExchange -match "Y") {
        while (!$ExSession) {
            Write-Output "`nWhat is the FQDN of the Exchange Server? (or type Exit)"
            $ExchangeServer = Read-Host -Prompt "FQDN"
            try {
                Write-Output "`nConnecting to Exchange Server Powershell Service"
                $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/
                Import-PSSession $ExSession
            }
            catch { Write-Output "! Unable to open Exchange session on: $ExchangeServer" }
            if ($ExchangeServer -match "Exit") { Exit; }
        }
    }
}

if ($isExo -match "Y" -or $isExchange -match "Y") {

    if (!$mailboxIntent) { 
        $intentResponse = ""
        Write-Output " "
        $intent = "`nDo you want to "
        if ($isExchange -match "Y") { 
            $intent += "`n`t[E]xport to PST and Delete "; $intentResponse += "E/" 
            $intent += "`n`t[M]ove to another database "; $intentResponse += "M/"
            $intent += "`n`t[P]ST Export Only "; $intentResponse += "P/"
        }
        else { $intent += "`n`t[C]onvert to Shared Mailbox "; $intentResponse += "C/" } 
        $intent += "`n`t[D]elete the mailbox, or "; $intentResponse += "D/"
        $intent += "`n`tMake [n]o changes to the target mailbox?`n"; $intentResponse += "N" 

        Write-Output "`n`n$intent"
        $mailboxIntent = Read-Host -Prompt $intentResponse; while ($intentResponse -notmatch $mailboxIntent) { $mailboxIntent = Read-Host $intentResponse } 
    }

    if (!$isRedirect) { Write-Output "`nDo you want to REDIRECT future emails set to target user to another email?"; $isRedirect = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isRedirect) { $isRedirect = Read-Host "Y/N" } }
    
    if ($isRedirect -match "Y"){
        While ($isRedirectEmail -notmatch "Y") {
            Write-Output "Please enter the Email for where emails intended for target user should be delivered"; $redirectEmail = Read-Host -Prompt 'Email'
            if (!$isRedirectEmail) { Write-Output "Please confirm redirect recipient email: $($redirectEmail)"; $isRedirectEmail = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isRedirectEmail) { $isRedirectEmail = Read-Host "Y/N" } }
        }
    }

    if ("E","D" -notcontains $mailboxIntent){
        if ($isRedirect -match "N") {
            if (!$isForward) { Write-Output "`nDo you want to FORWARD target user emails to another email?"; $isForward = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isForward) { $isForward = Read-Host "Y/N" } }
            if ($isForward -match "Y"){
               if (!$isLeaveCopy) { Write-Output "Do you want to LEAVE A COPY in target users mailbox?"; $isLeaveCopy = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isLeaveCopy) { $isLeaveCopy = Read-Host "Y/N" } }
                While ($isForwardEmail -notmatch "Y") {
                    $isForwardEmail = $null
                    Write-Output "Please enter the Email for where emails intended for target user should be forwarded"; $ForwardEmail = Read-Host -Prompt 'Email'
                    if (!$isForwardEmail) { Write-Output "Please confirm redirect recipient email: $($forwardEmail)"; $isForwardEmail = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isForwardEmail) { $isForwardEmail = Read-Host "Y/N" } }
                }
            }
        }

        if (!$isDelegated) { Write-Output "`nDo you want to DELEGATE ACCESS to target users mailbox to another user?"; $isDelegated = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isDelegated) { $isDelegated = Read-Host "Y/N" } }
        if ($isDelegated -match "Y" -and $isAADSync -match "Y"){
            While ($isDelegateUser -notmatch "Y"){
                $isDelegateUser = $null
                Write-Output "Select the delegate user from the list.."; $delegateUser = Get-ADUser -Filter * | select Enabled,Name,SamAccountName,UserPrincipalName | Out-GridView -PassThru
                if (!$isDelegateUser) { Write-Output "Please confirm delegated access user: $($delegateUser.Name)"; $isDelegateUser = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isDelegateUser) { $isDelegateUser = Read-Host "Y/N" } }
            }
            $delegateUPN = $delegateUser.UserPrincipalName
        }
        elseif ($isDelegated -match "Y" -and $isAADSync -match "N"){
            While ($isDelegateUPN -notmatch "Y") {
                $isDelegateUPN = $null
                Write-Output "Select the delegate user from the list.."; $delegateUPN = (Get-MsolUser | Out-GridView -PassThru).UserPrincipalName
                #Write-Output "Please enter the User Principal Name (UPN) or Email for the delegate user"; $delegateUPN = Read-Host -Prompt 'UPN/Email'
                if (!$isDelegateUPN) { Write-Output "Please confirm redirect recipient email: $($delegateUPN)"; $isDelegateUPN = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isDelegateUPN) { $isDelegateUPN = Read-Host "Y/N" } }
            }
        }

        if (!$isRestricted) { Write-Output "`nDo you want to RESTRICT DELIVERY of new emails to the target users mailbox?"; $isRestricted = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isRestricted) { $isRestricted = Read-Host "Y/N" } }
        if (!$isHidden) { Write-Output "`nDo you want to hide this user from the GAL?"; $isHidden = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isHidden) { $isHidden = Read-Host "Y/N" } }
    }
    elseif ("E","P" -contains $mailboxIntent){
        While ($isPath -notmatch "Y") { 
            $isPath = $null
            Write-Output "`nPlease enter the path to export the mailbox to, example \\servername\share\username.pst"; $path = Read-Host -Prompt "Path"
            if (!$isPath) { Write-Output "Please confirm path: $($path))"; $isPath = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isPath) { $isPath = Read-Host "Y/N" } }
            if (!(Test-Path $path)) { Write-Output "Unable to verify path.  Please create appropriate folder/share structure or provide a new path."; $isPath = "N" }
        }
    }
      

    
    if (!$isAutoResponse) { Write-Output "`nDo you want to set an AUTOMATED RESPONSE for emails sent to this user?"; $isAutoResponse = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isAutoResponse) { $isAutoResponse = Read-Host "Y/N" } }
    
    if ($isAutoResponse -match "Y"){ 
        if ("E","D" -notcontains $mailboxIntent) {
            if (!$ReponderType) { Write-Output "Do you want to deliver auto-response via NDR using [T]ransport Rule or from the target users [M]ailbox using out-of-office reply?"; $ResponderType = Read-Host "T/M"; while ("T","M" -notcontains $ResponderType) { $ResponderType = Read-Host "T/M" } }
        }
        else {
            $ResponderType = "T"
        }
        While ($isAutoResponseMessage -notmatch "Y"){
            $isAutoResponseMessage = $null
            Write-Output "Please enter text for auto-response.  Example: <User> is no longer with <Company>.  Please contact <Manager> by phone <000-000-0000> or email <name@domain.com>" 
            $AutoResponseMessage = Read-Host -Prompt 'Message'
            if (!$isAutoResponseMessage) { Write-Output "Please confirm the auto-response message: `n$($AutoResponseMessage)`n`n"; $isAutoResponseMessage = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isAutoResponseMessage) { $isAutoResponseMessage = Read-Host "Y/N" } }
        }
    }

    
    
}



########### DO STUFF ##########



if ($isDomain -match "Y") { 
    Write-Output "`nResetting Active Directory Password..."
    #Reset-AD-Password ($upn) 
    
}

if ($isOffice365 -match "Y" -and $isAADSync -match "Y") { 
    Write-Output "`nPerforming Azure AD Sync"
    #Sync-Password ($AADConnectServer) 
}

if ($isOffice365 -match "Y" -and $isAADSync -match "N") { 
    Write-Output "`nResetting Azure AD Password..."
    #Reset-AAD-Password ($upn) 
}

if ($isOffice365 -match "Y") { 
    Write-Output "`nBlocking User Access to Office 365..."
    #BlockUser($upn) 
    Write-Output "Completed"
}
if ($isExchange -match "Y" -or $isEXO -match "Y"){
    Write-Output "`nDisabling user connections (OWA, ActiveSync, MAPI, IMAP & POP)..."
    #DiableUserConnections($upn)
    Write-Output "Completed"
    
    Write-Output "`nRemoving User's devices (this does not wipe)..."
    #$userDevices = GetUserDevices($upn)
    #$numDevicesRemoved = RemoveDevices($userDevices)
    Write-Output "$($numDevicesRemoved) devices have been removed."

    # If Possible, Convert to Shared Mailbox
    
    if ($isRedirect -match "Y"){
        Write-Output "`nCreating Transport Rule to redirect all future emails from target to $redirectEmail"
        #RedirectEmail($upn,$redirectEmail)
        Write-Output "Completed"
    }

    if ($isForward -match "Y"){
        Write-Output "`nForwarding Emails from $upn to $ForwardEmail"
        #if ($isLeaveCopy -match "Y"){ DeliverAndForwardEmail ($upn,$ForwardEmail) }
        #else { FordwardEmail ($upn,$ForwardEmail) }
        Write-Output "Completed"
    }

    if ($isDelegated -match "Y"){
        Write-Output "`nSetting Mailbox Permissions for Delegate user $delegateUPN"
        #Add-MailboxDelegate ($upn, $delegateUPN)
        Write-Output "Completed"
    }

    if ($isRestricted -match "Y"){
        Write-Output "`nRestricting Delivery of New Emails to Target User Mailbox"
        #RestrictEmail ($upn)
        Write-Output "Completed"
    }

    if ($isAutoResponse -match "Y"){
        Write-Output "`nCreating auto-responder"
        #if ($ResponderType -match "T"){ Set-AutoResponse-Transport ($upn,$AutoResponseMessage) }
        #else { Set-AutoResponse-Mailbox ($upn, $AutoResponseMessage) }
        Write-Output "Completed"
    }

    if ($isHidden -match "Y"){
        Write-Output "`nHiding Target User from GAL"
        #HiddenEmail ($upn)
        Write-Output "Completed"
    }

    if ($mailboxIntent -match "E") {
        Write-Output "`nExporting Mailbox to $path"
        #ExportEmail ($upn,$path)
        Write-Output "Export request has been submitted."
        #while (!(Test-Path $path)){ 
            #delete
        #}
            
    }

    if ($mailboxIntent -match "D") {
        #delete
    }

    if ($mailboxIntent -match "C") {
        # Convert to Shared Mailbox
        if ($isSPO, $isODB -notcontains "Y") {
            # Remove License
            # Remove User
        }

    }
}

if ($isExchange -match "Y") { 


}
if ($isEXO -match "Y") { 


}
if ($isSPO -match "Y") { Write-Output "SharePoint Offboarding Processes are not supported at this time.  Please perform these tasks manually. " }
if ($isODB -match "Y") { Write-Output "OneDrive for Business Offboarding Processes are not supported at this time.  Please perform these tasks manually. " }

if ($ExSession -match "Open") { Remove-PSSession $ExSession }
#if ($ExoSession -match "Open") { Remove-PSSession $ExoSession }
if ($EopSession -match "Open") { Remove-PSSession $EopSession }

### TO DO ####

# 1. Add support for ODB/SPO/SFB
# 2. If target user is a Manager, update all users with Manager = TargetUser to New Manager
# 3. Export Mailbox to PST 
# 4. Delete Mailbox
# 5. Remove License
# 6. Disable or Delete User
# 7. Delegation of Home directory

# Move Target Users Home Share from \\SERVER\Users to \\SERVER\Users\EOE
# Remove all non-system/administor permissions
# Create Security Group 'EOE User - Name - Personal Folder Access'
# Add Security Group with Full Access permissions to Target User's Personal Folder
# Add delegated users to Security Group

  
