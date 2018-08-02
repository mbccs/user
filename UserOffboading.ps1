<#

.SYNOPSIS
  User Offboarding 
.DESCRIPTION
  Assists technical resource with properly offboarding a user
.PARAMETER
    
.INPUTS
  
.OUTPUTS
  Log file stored in C:\Temp
.NOTES
  Version:        1.0
  Author:         Jigar Shah
  Creation Date:  August 1, 2018
  Purpose/Change: Initial script development
 
.EXAMPLE
  

#>
###########################
##### PARAMTERS #####
param(
    [string]$adminCredential = $null,
    [string]$isDomain = $null,
    [string]$isCloudOnly = $null,
    [string]$isTargetUser = $null,
    [string]$isUPN = $null,
    [string]$isOffice365 = $null,
    [string]$isAADSync = $null,
    [string]$isAADConnectServer = $null,
    [string]$isEXO = $null,
    [string]$isExchange = $null,
    [string]$ExchangeServer = $null,
    [string]$isSPO = $null,
    [string]$isODB = $null,
    [string]$mailboxIntent = $null,
    [string]$intentResponse = $null,
    [string]$isRedirect = $null,
    [string]$isRedirectEmail = $null,
    [string]$redirectEmail = $null,
    [string]$isForward = $null,
    [string]$isForwardEmail = $null,
    [string]$ForwardEmail = $null,
    [string]$isLeaveCopy = $null,
    [string]$isDelegated = $null,
    [string]$isDelegateUser = $null,
    [string]$delegateUser = $null,
    [string]$isDelegateUPN = $null,
    [string]$delegateUPN = $null,
    [string]$isRestricted = $null,
    [string]$isHidden = $null,
    [string]$convertSharedAction = $null,
    [string]$isLicenseRemoved = $null,
    [string]$isPath = $null,
    [string]$isMailboxDB = $null,
    [string]$isAutoResponse = $null,
    [string]$ReponderType = $null,
    [string]$isAutoResponseMessage = $null,
    [string]$AutoResponseMessage = $null,
    [string]$homeDirectory = $null,
    [string]$homeDirIntent = $null,
    [string]$homeDirIntentResponse = $null,
    [string]$isHomeDirDelegate = $null,
    [string]$isHomePath = $null,
    [string]$homeDirDelegate = $null,
    [string]$delegateHomeDir = $null,
    [string]$homeDirPath = $null,
    [string]$profilePath = $null,
    [string]$profileIntent = $null,
    [string]$profileIntentResponse = $null,
    [string]$profileDirectory = $null,
    [string]$isProfileDelegate = $null,
    [string]$isProfilePath = $null,
    [string]$ProfileDelegate = $null,
    [string]$delegateProfileDir = $null,
    [string]$destinationHome = $null,
    [string]$destinationProfile = $null,
    [string]$destHomePath = $null,
    [string]$destProfilePath = $null,
    [string]$userObjectIntent = $null
)

#####################
##### FUNCTIONS #####



function ConnectMsol {
    Import-Module MSOnline -ErrorAction SilentlyContinue
    if (!(Get-Module MSOnline)) {
        Write-Host "MSOnline Module Not Found. Starting Download..."
        $source = "https://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_64.msi"
        $destination = "C:\temp"
        $filename = $source.Split("/")[-1]

        try {
            Start-BitsTransfer $source -Destination $destination
            Write-Host "Completed download from $($source) and saved to $($destination)."
        }
        catch {
            Write-Host "Unable to download file from $($source): $($_.Exception.Message)"
            Write-Host "Please download and install the Microsoft Online Services Sign-In Assistant for IT Professionals manually. Exiting."
            exit;
        }

        try {
            if (Test-Path -Path "$($destination)`\$($filename)") {
                $ArgsInstallation = @(
                    "/i"
                    ('"{0}"' -f "$($destination)`\$($filename)")
                    "ADDLOCAL=ALL"
                    "ACCEPT=YES"
                    "/qb"
                    "/norestart"
                )
                Start-Process msiexec.exe -ArgumentList $ArgsInstallation -Wait -Verbose
                Write-Host "Completed installation of $($filename)"

                Write-Host "Installing NuGet Package Provider..."
                Start-Process Install-PackageProvider -Name NuGet -Force -Wait
                Write-Host "Installed NuGet Package Provider."
                Write-Host "Installing MSOnline Module..."
                Install-Module MSOnline -Confirm:$false -Force -Wait
                Write-Host "Installed MSOnline Module."
            }
        }
        catch {
            Write-Host "Unable to install $($filename): $($_.Exception.Message)"
            Write-Host "Please complete installation manually. Exiting."
            exit;
        }
    }
    else { Write-Host "MSOnline Module is installed." }

    Import-Module MSOnline

    Write-Host "`nPlease submit credentials to connect to Office 365."
    $adminCredential = Get-Credential

    Write-Host "`nConnecting to AAD Powershell Service"
    Connect-MsolService -Credential $adminCredential -ErrorAction SilentlyContinue 

    return $adminCredential
}

function Reset-AD-Password($upn) {
    $Password = ([char[]]([char]33..[char]95) + ([char[]]([char]97..[char]126)) + 0..9 | sort {Get-Random})[0..14] -join ''
    $SecPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

    $user = get-aduser -filter { UserPrincipalName -like $upn } 
    $user | Set-ADAccountPassword -Reset -NewPassword $SecPassword
    $user | Set-ADUser -ChangePasswordAtLogon $true   

    Write-Host "The password for the user account $upn has been set to be $Password.  Make sure you record this and share with the user, or be ready to reset the password again.  They will have to reset their password on the next logon."
}

function Reset-AAD-Password($upn) {
    $Password = ([char[]]([char]33..[char]95) + ([char[]]([char]97..[char]126)) + 0..9 | sort {Get-Random})[0..14] -join ''
    $SecPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

    Set-MsolUserPassword –UserPrincipalName $upn –NewPassword $secPassword -ForceChangePassword $True
    Write-Host "The password for the account $upn has been set to be $Password. Make sure you record this and share with the user, or be ready to reset the password again. They will have to reset their password on the next logon."

    Set-MsolUser -UserPrincipalName $upn -StrongPasswordRequired $True
    Write-Host "This user's account has also been set to require a strong password."

}

function Sync-Password($adsync) {
    try {
        Invoke-Command -ComputerName $adsync -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta } 
    }
    catch {
        Write-Host "Unable to complete synchronization request.  Please perform sync manually. "
    }        
}

Function BlockUser($upn) {
    Set-MsolUser –UserPrincipalName $upn –blockcredential $true
}

Function DiableUserConnections($upn){
    Set-CASMailbox $upn -OWAEnabled $False -ActiveSyncEnabled $False –MAPIEnabled:$false -IMAPEnabled:$false -PopEnabled:$false 
}

Function GetUserDevices($upn){  
    $userMobileDevice = Get-MobileDevice -Mailbox $upn
    return $userMobileDevice
}

Function RemoveDevices($userMobileDevice){
    
    if (!$userMobileDevice) { return 0; }
    else{
        $i = 0
        while ($i -lt $userMobileDevice.length){ Remove-MobileDevice -Identity $userMobileDevice[$i]; $i++ }
        return $i
    }
}

Function RedirectEmail($redirectFrom,$redirectTo){
    $currentDate = (Get-Date).ToString("yyyyMMdd")
    $ruleName = "EOE Forward created on $currentDate from $redirectFrom to $redirectTo"
    New-TransportRule -Name $ruleName -SentTo $redirectFrom -RedirectMessageTo $redirectTo
}

function Remove-MailboxDelegates($upn) {
    #Write-Host "`nRemoving Mailbox Delegate Permissions for the target user $upn."

    $mailboxDelegates = Get-MailboxPermission -Identity $upn | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
    Get-MailboxPermission -Identity $upn | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
    
    foreach ($delegate in $mailboxDelegates) {
        Remove-MailboxPermission -Identity $upn -User $delegate.User -AccessRights $delegate.AccessRights -InheritanceType All -Confirm:$false
    }

    #TO DO: Need to figure out how to check delegate permissions set on a all the folders for the user, then remove them. Looks to be a user-only cmdlet permission set
    #$mailboxFolders = Get-MailboxFolder -Identity admin -Recurse
    #foreach ($folder in $mailboxFolders) 
    #{
    #    $thisUpnFolder = $upn + ":\" + $folder.FolderPath
    #    Get-MailboxFolderPermission -Identity $thisUpnFolder | Where-Object {($_.AccessRights -ne "None")}
    #Remove-MailboxFolderPermission: https://technet.microsoft.com/en-us/library/dd351181(v=exchg.160).aspx
}

function Add-MailboxDelegate($upn,$delegateUser){
    Add-MailboxPermission -Identity $upn -User $delegateUser -AccessRights FullAccess -InheritanceType All
    Write-Host "$delegateUser has been given Mailbox Delegate Permissions for the target user $upn."
    Get-MailboxPermission -Identity $upn | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
}

Function Add-SMTP($upn,$smtp){
    Set-ADUser -Identity $upn -Add @{Proxyaddresses="SMTP:"+$smtp}
}

Function Remove-SMTP($upn,$smtp){
    Set-ADUser -Identity $upn -Remove @{Proxyaddresses="SMTP:"+$smtp}
}

Function RestrictEmail ($upn){
    Set-Mailbox -Identity $upn -AcceptMessagesOnlyFrom @{add="Administrator"}
}

Function Set-AutoResponse-Transport ($upn,$message){
    $currentDate = (Get-Date).ToString("yyyyMMdd")
    $ruleName = "EOE Auto-responder created on $currentDate for $upn"
    New-TransportRule -Name $ruleName -SentTo $upn -RejectMessageReasonText $message
}

Function Set-AutoResponse-Mailbox ($upn,$message){
    Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -InternalMessage $message -ExternalMessage $message
}

Function HiddenEmail ($upn){
    Set-Mailbox -Identity $upn -HiddenFromAddressListsEnabled $true
}

Function ForwardEmail($forwardFrom,$forwardTo){
    Set-Mailbox -Identity $forwardFrom -ForwardingSMTPAddress $forwardTo
}

Function DeliverAndForwardEmail ($forwardFrom,$forwardTo){
    Set-Mailbox -Identity $forwardFrom -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $forwardTo
}

Function ExportArchiveMailbox ($upn,$path){
    if (!(Get-Mailbox jshah@mbccs.com -Archive -ErrorAction SilentlyContinue)){
        Write-Host "Mailbox Archive is not enabled."
    else 
        try {
            New-MailboxExportRequest -Name "EOE_$upn" -Mailbox $upn -IsArchive -FilePath $path+"InPlaceHold_"+$filename+"_archive.pst"
            While (!(Get-MailboxExportRequest -Mailbox $upn -Status Completed)) { Start-Sleep -s 300 }
        }
        catch { Write-Host "Please confirm logged in user has Mailbox Import Export role assigned and try again." }
    }
}

Function ExportMailbox ($upn,$path){
    $filename = $upn.split("@")[0]
    if ($path[-1] -ne "`\") { $path += "`\" }
    try {
        New-MailboxExportRequest -Name "EOE_$upn" -Mailbox $upn -FilePath $path+"InPlaceHold_"+$filename+".pst"
        While (!(Get-MailboxExportRequest -Mailbox $upn -Status Completed)) { Start-Sleep -s 300 }
        ExportArchiveMailbox ($upn,$path)
    }
    catch { Write-Host "Please confirm logged in user has Mailbox Import Export role assigned and try again." }
        
}

Function DisableMailbox ($upn){
    Disable-Mailbox $upn
}

Function DeleteMailbox ($upn){
    $DisplayName = (Get-Mailbox $upn).Name
    Disable-Mailbox $upn
    $guid = Get-MailboxDatabase | Get-MailboxStatistics | where { $_.DisplayName -match $name -and $_.DisconnectDate -ne $null } | select MailboxGuid
    Get-MaiboxDatabase | Remove-Mailbox -StoreMailboxIdentity $guid
}

Function ConvertToSharedMailbox ($upn){
    Set-Mailbox $upn -Type Shared
}

Function RemoveLicenses ($upn){
    $licenseObj = Get-MsolAccountSku
    $license = $licenseObj.AccountSkuId
    Set-MsolUserLicese -UserPrincipalName $upn -RemoveLicenses $license
}

Function RemoveMsolUser ($upn){
    Remove-MsolUser -UserPrincipalName $upn
}

Function PurgeDirectory ($path) {
    takeown /F $path /A /R /D y
    icacls $path /t /grant administrators:f  
    Remove-Item $path -Recurse -Force 
}

Function ArchiveDirectory ($source,$destination){
    try {
        Add-Type -AssemblyName "system.io.compression.filesystem"
        [io.compression.zipfile]::CreateFromDirectory($source,$destination)
        Write-Host "$source has been compressed and placed in $destination"
        if (Test-Path $destination){
            PurgeDirectory($source)
        }
    }
    catch {
        Write-Host "Unable to compress $source"
    }
}

Function MoveDirectory ($source,$destination){
    try {
        &robocopy $source $destination /S /E R:1 /W:1 /MOV
    }
    catch {
        Write-Host "Unable to move $source to $destination"
   }

}

Function DisableUserObject ($upn) {
    try { get-aduser -filter { UserPrincipalName -like $upn } | Disable-ADAccount }
    catch { Write-Host "Unable to disable user object" }
}

Function DeleteUserObject ($upn){
    DisableUserObject ($upn)
    try { Get-ADUser -Filter { UserPrincipalName -like $upn } | Remove-ADUser }
    catch { Write-Host "Unable to delete user object" }
}

Function RemoveUserObject ($upn) {
    try{
        $DGs = Get-DistributionGroup -ErrorAction SilentlyContinue
        $PGs = get-aduser -filter { UserPrincipalName -like $upn } | Get-ADPrincipalGroupMembership -ErrorAction SilentlyContinue

        if ($DGs) { foreach ($DG in $DGs) { Remove-DistributionGroupMember -Identity $dg -Member $upn -ErrorAction SilentlyContinue -Confirm:$false } }
        if ($PGs) { Remove-ADPrincipalGroupMembership -Identity $upn.split("@")[0] -MemberOf $PGs -Confirm:$false }
    }
    catch{
        Write-Host "Unable to remove user from any groups"
    }
}




###########################

if (!$isDomain) { Write-Output "`nIs the offboarding target user on Active Directory?"; $isDomain = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isDomain) { $isDomain = Read-Host "Y/N" } }
if ($isDomain -match "N") { if (!$isCloudOnly) { Write-Output "`nIs the offboarding target user a cloud-only (Azure AD) user?"; $isCloudOnly = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isCloudOnly) { $isCloudOnly = Read-Host "Y/N" } } }

if ($isDomain -match "Y") {
    try { Import-Module ActiveDirectory }
    catch { Write-Output "User Offboarding Script must be executed from a server with ActiveDirectory module available.  Please re-run from a Domain Controller or install RSAT Tools. "; exit; }
    
    while ($isTargetUser -notmatch "Y") {
        $isTargetUser = $Null
        Write-Output "`nSelect the target from the list.."; $targetUser = Get-ADUser -Filter * -Properties Enabled,Name,SamAccountName,UserPrincipalName,HomeDirectory,ProfilePath | select Enabled,Name,SamAccountName,UserPrincipalName,HomeDirectory,ProfilePath | Out-GridView -PassThru
        if (!$isTargetUser) { Write-Output "Please confirm target user: $($targetUser.Name)"; $isTargetUser = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isTargetUser) { $isTargetUser = Read-Host "Y/N" } }
        if (!$TargetUser) { Write-Output "Target user was not selected.  Please select a user. "; $isTargetUser = $null }
    }
    $userName = $targetUser.SamAccountName
    $upn = $targetUser.UserPrincipalName
    $homeDirectory = $targetUser.HomeDirectory
    if ($homeDirectory){
        if (!(Test-Path $homeDirectory)){ Write-Output "`nUnable to confirm HOME directory.  Continuing.. "; $homeDirectory = $null }
    }
    $profileDirectory = $targetUser.ProfilePath
    if ($profileDirectory) {
        if (!(Test-Path $profiledirectory)) { Write-Output "`nUnable to confirm PROFILE directory. Continuing.. "; $profileDirectory = $null }
    }
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
        if (!$targetUser) { Write-Output "Target user was not selected.  Please select a user. "; $isTargetUser = $null }
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

Write-Output ""
$transcriptpath = "C:\Temp\User_Offboarding_" + $userName + "_" +(Get-Date).ToString('yyyy-MM-dd') + ".txt"
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

if ($isExo,$isExchange -contains "Y") {

    if (!$mailboxIntent) { 
        $intentResponse = ""
        Write-Output " "
        $intent = "`nWhat do you want to do with the target users Mailbox: "
        if ($isExchange -match "Y") { 
            $intent += "`n`t[E]xport to PST and Disable Mailbox (mailbox will be removed on retention period expiry) "; $intentResponse += "E/" 
            #$intent += "`n`t[M]ove to another database "; $intentResponse += "M/"
            $intent += "`n`t[P]ST Export Only "; $intentResponse += "P/"
        }
        else { $intent += "`n`t[C]onvert to Shared Mailbox "; $intentResponse += "C/" } 
        $intent += "`n`t[S]oft delete (disable) the mailbox "; $intentResponse += "S/"
        $intent += "`n`t[D]elete the mailbox immediately, or "; $intentResponse +="D/" 
        $intent += "`n`t[N]one of the above are required`n"; $intentResponse += "N" 

        Write-Output "`n`n$intent"
        $mailboxIntent = Read-Host -Prompt $intentResponse; while ($intentResponse -notmatch $mailboxIntent) { $mailboxIntent = Read-Host $intentResponse } 
    }

    if (!$isRedirect) { Write-Output "`nDo you want to REDIRECT future emails set to target user to another email? `nNote: This option will create a transport rule which will forward all emails sent to $upn to a different email.  No email messages will be delivered to a mailbox. "; $isRedirect = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isRedirect) { $isRedirect = Read-Host "Y/N" } }
    
    if ($isRedirect -match "Y"){
        While ($isRedirectEmail -notmatch "Y") {
            $isRedirectEmail = $null
            Write-Output "Please enter the Email for where emails intended for target user should be delivered"; $redirectEmail = Read-Host -Prompt 'Email'
            if (!$isRedirectEmail) { Write-Output "Please confirm redirect recipient email: $($redirectEmail)"; $isRedirectEmail = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isRedirectEmail) { $isRedirectEmail = Read-Host "Y/N" } }
        }
    }

    if ("E","D","S" -notcontains $mailboxIntent){
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

        if (!$isDelegated) { Write-Output "`nDo you want to DELEGATE ACCESS to target mailbox to another user?"; $isDelegated = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isDelegated) { $isDelegated = Read-Host "Y/N" } }
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

        if (!$isRestricted) { Write-Output "`nDo you want to RESTRICT DELIVERY of new emails to target mailbox?"; $isRestricted = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isRestricted) { $isRestricted = Read-Host "Y/N" } }
        if (!$isHidden) { Write-Output "`nDo you want to hide this user from the GAL?"; $isHidden = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isHidden) { $isHidden = Read-Host "Y/N" } }
    }
    
    if ("C" -contains $mailboxIntent) {
        if ((((Get-Mailbox $upn).ProhibitSendQuota).Split(" ")[0]) -gt "50"){
            if (!$convertSharedAction) { Write-Output "Mailbox cannot be converted to shared mailbox due to size.  Shared Mailbox size limit is 50GB.  Would you like to export to PST?"; $convertSharedAction = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $convertSharedAction) { $convertSharedAction = Read-Host "Y/N" } }
        }
        if (Get-Mailbox $upn -Archive -ErrorAction SilentlyContinue) {
            if (!$convertSharedAction) {  Write-Output "`nTarget Mailbox has archiving enabled.  Please confirm if you would like to [k]eep the shared mailbox licensed to prevent data loss, or if you would like to [e]xport the archive mailbox to a PST and proceed with conversion and license removal."; $convertSharedAction = Read-Host -Prompt "K/E"; while ("K","E" -notcontains $convertSharedAction) { $convertSharedAction = Read-Host "K/E" } }
        }
        if (("K","N" -contains $convertSharedAction) -or ($isODB,$isSPO -contains "Y")) { $isLicenseRemoved = "N" }
        if (!$isLicenseRemoved) { Write-Output "`nDo you want to remove target user from Office 365/Azure AD and return assigned licenses?"; $isLicenseRemoved = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isLicenseRemoved) { $isLicenseRemoved = Read-Host "Y/N" } }
    }

    if (("E","P" -contains $mailboxIntent) -or ("E","Y" -contains $convertSharedAction)){
        While ($isPath -notmatch "Y") { 
            Write-Output "`nPlease enter the path to export the mailbox and/or archive mailbox to, example \\servername\share\"; $path = Read-Host -Prompt "Path"
            if (!$isPath) { Write-Output "Please confirm path: $($path)"; $isPath = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isPath) { $isPath = Read-Host "Y/N" } }
            if ($isPath -match "Y") { if (!(Test-Path $path)) { Write-Output "Unable to verify path.  Please create appropriate folder/share structure or provide a new path."; $isPath = $null } }
            else { $isPath = $null }
        }
    }
    
    if ("M" -contains $mailboxIntent){
        while ($isMailboxDB -notmatch "Y"){

        }
    }  

    
    if (!$isAutoResponse) { Write-Output "`nDo you want to set an AUTOMATED RESPONSE for emails sent to this user?"; $isAutoResponse = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isAutoResponse) { $isAutoResponse = Read-Host "Y/N" } }
    
    if ($isAutoResponse -match "Y"){ 
        if ("D","E","S" -notcontains $mailboxIntent) {
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

if ($isDomain -match "Y"){
    if ($homeDirectory){
        if (!$homeDirIntent) {
            $homeDirIntentResponse = ""
            Write-Output " "
            $intent = "`nWhat do you want to do with the target users HOME Directory: "
            $intent += "`n`t[P]ermanently delete the directory "; $homeDirIntentResponse += "P/"
            $intent += "`n`t[T]ransfer the directory another user as a subfolder "; $homeDirIntentResponse +="T/" 
            $intent += "`n`t[M]ove the directory to a location on a shared network drive, or "; $homeDirIntentResponse +="M/" 
            $intent += "`n`t[Z]ip the directory for archiving`n"; $homeDirIntentResponse += "Z" 

            Write-Output "`n`n$intent"
            $homeDirIntent = Read-Host -Prompt $homeDirIntentResponse; while ($homeDirIntentResponse -notmatch $homeDirIntent) { $homeDirIntent = Read-Host $homeDirIntentResponse } 
        } 

        switch ($homeDirIntent){
            "T"{
                While ($isHomeDirDelegate -notmatch "Y"){
                    Write-Output "`nSelect the delegate user from the list.."; $homeDirDelegate = Get-ADUser -Filter * -Properties Enabled,Name,SamAccountName,UserPrincipalName,HomeDirectory,ProfilePath | select Enabled,Name,SamAccountName,UserPrincipalName,HomeDirectory,ProfilePath | Out-GridView -PassThru
                    if (!$isHomeDirDelegate) { Write-Output "`nPlease confirm delegated access user: $($homeDirDelegate.Name)"; $isHomeDirDelegate = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isHomeDirDelegate) { $isHomeDirDelegate = Read-Host "Y/N" } }
                    if (!$homeDirDelegate.HomeDirectory) { Write-Output "This user does not have a home directory defined.  Please select another user."; $isHomeDirDelegate = $null }
                    else { 
                        if (Test-Path $homeDirDelegate.HomeDirectory){
                            $destHomePath = Join-Path -Path $homeDirDelegate.HomeDirectory -ChildPath "$($username)_personal"
                            try { New-Item -ItemType Directory -Path $homeDirDelegate.HomeDirectory -Name "$($username)_personal" }
                            catch { Write-Output "`nUnable to create destination path. "; $isHomeDirDelegate = $null }
                        }
                        else { Write-Output "`nUnable to find destination folder, please confirm it exists and permissions allow active user to write data before trying again. "; $isHomeDirDelegate = $null }
                    }
                    if ($isHomeDirDelegate -match "N") { $isHomeDirDelegate = $null }
                }
                $delegateHomeDir = $homeDirDelegate.HomeDirectory
            }
            "M"{
                While ($isHomePath -notmatch "Y") { 
                    Write-Output "`nPlease enter the destination path to move the home directory to, example \\servername\share\eoeusers\"; $homeDirPath = Read-Host -Prompt "Path"
                    if (!$isHomePath) { Write-Output "Please confirm path: $($homeDirPath)"; $isHomePath = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isHomePath) { $isHomePath = Read-Host "Y/N" } }
                    if ($isHomePath -match "Y") { 
                        if (!(Test-Path $homeDirPath)) { Write-Output "Unable to verify path.  Please create appropriate folder/share structure or provide a new path."; $isHomePath = $null } 
                        else { 
                            $destHomePath = Join-Path -Path $homeDirPath -ChildPath "$($username)_personal"
                            if (Test-Path $destHomePath) { Write-Output "The destination path $destHomePath already exists. Please choose a different path to prevent overwrite."; $isHomePath = $null }
                            else {
                                try { New-Item -ItemType Directory -Path $homeDirPath -Name "$($username)_personal" }
                                catch { Write-Output "Unable to create destination path "; $isHomePath = $null }
                            }
                        }
                    }
                    else { $isHomePath = $null }
                }
            }
            "Z" {
                $destinationHome = $homeDirectory.Substring(0, $homeDirectory.LastIndexOf('\'))+"`\$username.zip"
            }
        }
    }
    if ($profileDirectory){
        if (!$profileIntent) {
            $profileIntentResponse = ""
            Write-Output " "
            $intent = "`nWhat do you want to do with the target users Roaming Profile: "
            $intent += "`n`t[P]ermanently delete the directory "; $profileIntentResponse += "P/"
            $intent += "`n`t[T]ransfer the directory another user as a subfolder of their Home Directory "; $profileIntentResponse +="T/" 
            $intent += "`n`t[M]ove the directory to a location on a shared network drive, or "; $profileIntentResponse +="D/" 
            $intent += "`n`t[Z]ip the directory for archiving`n"; $profileIntentResponse += "Z" 

            Write-Output "`n`n$intent"
            $profileIntent = Read-Host -Prompt $profileIntentResponse; while ($profileIntentResponse -notmatch $profileIntent) { $profileIntent = Read-Host $profileIntentResponse } 
        } 
        switch ($profileIntent){
            "T"{
                While ($isProfileDelegate -notmatch "Y"){
                    Write-Output "Select the delegate user from the list.."; $ProfileDelegate = Get-ADUser -Filter * -Properties Enabled,Name,SamAccountName,UserPrincipalName,HomeDirectory,ProfilePath | select Enabled,Name,SamAccountName,UserPrincipalName,HomeDirectory,ProfilePath | Out-GridView -PassThru
                    if (!$isProfileDelegate) { Write-Output "Please confirm delegated access user: $($ProfileDelegate.Name)"; $isProfileDelegate = Read-Host -Prompt "Y/N"; while ("Y", "N" -notcontains $isProfileDelegate) { $isProfileDelegate = Read-Host "Y/N" } }
                    if (!$ProfileDelegate.HomeDirectory) { Write-Output "This user does not have a home directory defined.  Please select another user."; $isProfileDelegate = $null }
                    else { 
                        if (Test-Path $profileDirDelegate.HomeDirectory){
                            $destProfilePath = Join-Path -Path $profileDirDelegate.HomeDirectory -ChildPath "$($username)_profile"
                            try { New-Item -ItemType Directory -Path $profileDirDelegate.HomeDirectory -Name "$($username)_profile" }
                            catch { Write-Output "`nUnable to create destination path "; $isProfileDelegate = $null }
                        }
                        else { Write-Output "`nUnable to find destination folder, please confirm it exists and permissions allow active user to write data before trying again. "; $isProfileDelegate = $null }
                    }
                    if ($isProfileDelegate -match "N") { $isProfileDelegate = $null }
                }
                $delegateProfileDir = $ProfileDelegate.HomeDirectory
            }
            "M"{
                While ($isProfilePath -notmatch "Y") { 
                    Write-Output "`nPlease enter the destination path to move the PROFILE directory to, example \\servername\share\eoeprofiles\"; $profileDirPath = Read-Host -Prompt "Path"
                    if (!$isProfilePath) { Write-Output "Please confirm path: $($profileDirPath)"; $isProfilePath = Read-Host -Prompt "Y/N"; while ("Y","N" -notcontains $isProfilePath) { $isProfilePath = Read-Host "Y/N" } }
                    if ($isProfilePath -match "Y") { 
                        if (!(Test-Path $profileDirPath)) { Write-Output "Unable to verify path.  Please create appropriate folder/share structure or provide a new path."; $isProfilePath = $null } 
                        else { 
                            $destProfilePath = Join-Path -Path $profileDirPath -ChildPath "$($username)_profile"
                            if (Test-Path $destProfilePath) { Write-Output "The destination path $destProfilePath already exists. Please choose a different path to prevent overwrite."; $isProfilePath = $null }
                            else {
                                try { New-Item -ItemType Directory -Path $profileDirPath -Name "$($username)_profile" }
                                catch { Write-Output "`nUnable to create destination path "; $isProfilePath = $null }
                            }
                        }
                    }
                    else { $isProfilePath = $null }
                }

            }
            "Z" {
                $destinationProfile = $profileDirectory.Substring(0, $profileDirectory.LastIndexOf('\'))+"`\$username.zip"
            }
        }
    }
    if (!$userObjectIntent) {
        $userObjectIntentResponse = ""
        Write-Output ""
        $intent = "`nWhat do you want to do with the User Object: "
        $intent += "`n`t[P]ermanently delete the user object"; $userObjectIntentResponse += "P/"
        $intent += "`n`t[D]isable the user object"; $userObjectIntentResponse += "D/"
        $intent += "`n`t[R]emove the user object from all security and distrubtion groups, and nothing else"; $userObjectIntentResponse += "R/"
        $intent += "`n`t[N]o changes to the user object"; $userObjectIntentResponse += "N/"
        Write-Output "`n`n$intent"
        $userObjectIntent = Read-Host -Prompt $userObjectIntentResponse; while ($userObjectIntentResponse -notmatch $userObjectIntent) { $userObjectIntent = Read-Host $profileIntentResponse }
    }
    switch ($userObjectIntent){
        #"P"{}
        #"D"{}
        "R"{}
        #"N"{}
    }
}



########### DO STUFF ##########

if ($isDomain -match "Y") { 
    Write-Output "`nResetting Active Directory Password..."
    #Reset-AD-Password ($upn) 
    
}

if ($isOffice365 -match "Y") { 
    if ($isAADSync -match "Y") {
        Write-Output "`nPerforming Azure AD Sync"
        #Sync-Password ($AADConnectServer) 
    }
    else {
        Write-Output "`nResetting Azure AD Password..."
        #Reset-AAD-Password ($upn) 
    }

    Write-Output "`nBlocking User Access to Office 365..."
    #BlockUser($upn) 
    Write-Output "Completed"
}
if ($isExchange,$isExo -contains "Y"){
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

    if ($AutoResponseMessage){
        if ($ResponderType -match "T"){ 
            Write-Output "`nCreating Transport Rule to auto-respond to emails sent to $upn"
            #Set-AutoResponse-Transport ($upn,$AutoResponseMessage)
        }
        else { 
            Write-Output "`nCreating Inbox Automated Reply for mailbox of $upn"
            #Set-AutoResponse-Mailbox ($upn, $AutoResponseMessage)
        }
        Write-Output "Completed"
    }

    if ($isHidden -match "Y"){
        Write-Output "`nHiding Target User from GAL"
        #HiddenEmail ($upn)
        Write-Output "Completed"
    }

    if ($path) {
        Write-Output "`nExporting Mailbox to $path.  Please wait until export has been completed..."
        #ExportEmail ($upn,$path)
        Write-Output "Completed"
    }

    if ("E","S" -contains $mailboxIntent) {
        Write-Output "`nDisabling mailbox for $upn"
        #DisableMailbox ($upn)
        Write-Output "Completed"
    }

    if ("D" -contains $mailboxIntent) {
        Write-Output "`nDisabling and Removing mailbox for $upn" 
        #DeleteMailbox ($upn)
        Write-Output "Completed"
    }

    if ("C" -contains $mailboxIntent) {
        # Convert to Shared Mailbox
        Write-Output "`nConverting target user mailbox to shared mailbox"
        #ConvertToSharedMailbox ($upn)
        Write-Output "Completed"
    }
}

if ($isSPO -match "Y") { Write-Output "SharePoint Offboarding Processes are not supported at this time.  Please perform these tasks manually. " }
if ($isODB -match "Y") { Write-Output "OneDrive for Business Offboarding Processes are not supported at this time.  Please perform these tasks manually. " }

if ($isLicenseRemoved -match "Y"){ 
    Write-Output "`nRemoving Office 365 Licenses from $upn"
    #RemoveLicenses ($upn)
    Write-Output "Completed"
    if ($isAADSync -notmatch "Y"){
        Write-Output "`nRemoving User from Azure AD/Office 365"
        #RemoveMsolUser ($upn)
        Write-Output "Completed"
    }
}

if ($ExSession -match "Open") { Remove-PSSession $ExSession }
if ($ExoSession -match "Open") { Remove-PSSession $ExoSession }
if ($EopSession -match "Open") { Remove-PSSession $EopSession }

switch ($homeDirIntent) {
    "P" {
        Write-Output "`nPermanently deleting $homeDirectory"
        #PurgeDirectory ($homeDirectory)
        Write-Output "Completed"
    }
    "T" {
        Write-Output "`nMoving user HOME directory to $destHomePath"
        #MoveDirectory ($homeDirectory,$destHomePath)
    }
    "M" {
        Write-Output "`nMoving user HOME directory to $destHomePath"
        #MoveDirectory ($homeDirectory,$destHomePath)
    }
    "Z" {
        Write-Output "`nArchiving $homeDirectory"
        #ArchiveDirectory ($homeDirectory, $destinationHome)
        
    }
}

switch ($profileIntent) {
    "P" {
        Write-Output "`nPermanently deleting $profileDirectory"
        #PurgeDirectory ($profileDirectory)
        Write-Output "Completed"
    }
    "T" {
        Write-Output "`nMoving user HOME directory to $destHomePath"
        #MoveDirectory ($homeDirectory,$destHomePath)
    }
    "M" {
        Write-Output "`nMoving user HOME directory to $destHomePath"
        #MoveDirectory ($homeDirectory,$destHomePath)
    }
    "Z" {
        Write-Output "`nArchiving $profileDirectory"
        #ArchiveDirectory ($profileDirectory, $destinationProfile)
    }
}

switch ($userObjectIntent){
    "P"{
        Write-Output "`nPermantently deleting user object"
        #DeleteUserObject ($upn)   
    }
    "D"{
        Write-Output "`nDisabling user object"
        #DisableUserObject ($upn) 
    }
    "R"{
        Write-Output "`nRemoving user from all security and distribution groups"
        #RemoveUserObject ($upn)
    }
}

### TO DO ####

# 1. Add support for ODB/SPO/SFB
# 2. If target user is a Manager, update all users with Manager = TargetUser to New Manager


  
