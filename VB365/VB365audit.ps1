<#
.SYNOPSIS
  VB365audit.ps1
.AUTHOR
  Commenge Damien, Axians Cloud Builder
.DESCRIPTION
  This script created an Audit Report for Veeam Backup for Microsoft 365
.NOTES
  Version:        1.1
  Author:         Julien Mousqueton, Computacenter 
  Creation Date:  18/07/2022
  Purpose/Change: Use $env:TEMP variable for the path to the temporary folder
                  Change VBM365 to VB365 
                  Optimize script
  
.COMMENT
   <N/A> is used for not available
#>

$date = (get-date -Format "dd_MM_yyyy_HH_mm")

#ReportPath
$ReportPath="$env:TEMP\VB365Audit\$date"
#Report file HTML path
$htmlReportPath = "$ReportPath\VeeamBackupMicrosoft365.html"


#Create directory
New-Item -ItemType Directory -Path "$ReportPath" -Force | Out-Null

$HTMLTitle = "VB365 report"
$HTMLCSS = @'
<style>
body{color:black;font-family:Vinci Sans Light;font-size:0.79em;line-height:1.25;margin:5;}
a{color:black;}
H1{color:white;font-family:Verdana;font-weight:bold;font-size:20pt;margin-bottom:50px;margin-top:40px;text-align:center;background-color:#005EB8;}
H2{color:#A20067;font-family:Verdana;font-size:16pt;margin-left:14px;text-align:left;}
H3{color:#005EB8;font-family:Verdana;font-size:13pt;margin-left:16px;}
H4{color:black;font-family:Verdana;font-size:11pt;margin-left:16px;}
table {border-collapse: collapse;margin-left:10px;border-radius:7px 7px 0px 0px;}
th, td {padding: 8px;text-align: left;border-bottom: 1px solid #ddd;}
th {background-color: #006400;color: white;}
td:first-child{font-weight:bold;}
tr:nth-child(even){background-color: #f2f2f2}
table.table2 td:first-child{background-color: #A20067;color: white}
</style>
'@

#Connect to VBO Server
Write-host "$(get-date -Format HH:mm) - Connecting to VB365 server"
try {
    Connect-VBOServer -ErrorAction Stop
    Write-host "$(get-date -Format HH:mm) - Connected to VB365 server"

}
catch [System.Management.Automation.RuntimeException]{
    Write-host "$(get-date -Format HH:mm) - Connexion is already done"
}
catch {
    Write-host "$(get-date -Format HH:mm) - $($_.Exception.message) " -ForegroundColor Red
    break
}

 <#
 .Synopsis
    Get configuration Summary from Veeam Microsoft 365 server
 .DESCRIPTION
    Get server name, OS, OS build and VBM365 version
 .EXAMPLE 
    Get-DCVBMSummary
 #>
function Get-DCVBMSummary
{
    Write-host "$(get-date -Format HH:mm) - VB365 Summary"

    $VBM365ServerName = $env:COMPUTERNAME
    $VBM365ServerOS = (Get-WmiObject win32_operatingsystem).caption
    $OSBuild = Get-ItemPropertyValue -path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion" -name 'UBR'
    #No way to find version with powershell but associated version with VBR is available here : https://www.veeam.com/kb4106 
    $VBRbuild = (Get-Module Veeam.Archiver.PowerShell).nestedmodules.version.tostring()
    switch ($VBRbuild)
    {
        "10.0.2.1061"   {$VBM365Build = "5.0.0.1061"}
        "10.0.2.1063"   {$VBM365Build = "5.0.0.1063"}
        "10.0.2.1070"   {$VBM365Build = "5.0.0.1070"}
        "10.0.3.179"    {$VBM365Build = "5.0.1.179"}
        "10.0.3.207"    {$VBM365Build = "5.0.1.207"}
        "10.0.3.225"    {$VBM365Build = "5.0.1.225"}
        "10.0.3.252"    {$VBM365Build = "5.0.1.252"}
        "10.0.4.22"     {$VBM365Build = "5.0.2.22"}
        "10.0.4.42"     {$VBM365Build = "5.0.2.42"}
        "10.0.5.1033"   {$VBM365Build = "5.0.3.1033"}
        "10.0.5.1035"   {$VBM365Build = "5.0.3.1035"}
        "10.0.5.1051"   {$VBM365Build = "5.0.3.1051"}
        "10.0.5.1060"   {$VBM365Build = "5.0.3.1060"}
        "10.0.5.1063"   {$VBM365Build = "5.0.3.1063"}
        "11.1.0.367"    {$VBM365Build = "6.0.0.367"}
        "11.1.0.379"    {$VBM365Build = "6.0.0.379"}
        "11.1.0.385"    {$VBM365Build = "6.0.0.385"}
        Default         {$VBM365Build = "Unknown Value, script update is necessary"}
    }
    
    [PScustomObject]@{
        Name = $VBM365ServerName
        OS = $VBM365ServerOS
        OSBuild = $OSBuild
        VBM365Version = $VBM365Build
    }
}

 <#
 .Synopsis
    Get configuration about organizations
 .DESCRIPTION
    Get organization name, account used, type (on premise, hybride, O365), service (exchange, sharepoint), region, authentication (basic, modern with legacy protocol, modern), auxiliar backup account/application number
 .EXAMPLE 
    Get-DCVBMOrganization
 #>
function Get-DCVBMOrganization
{
    Write-host "$(get-date -Format HH:mm) - VB365 Organization"

    $OrgName = (Get-VBOOrganization).OfficeName
    $OrgAccount = (Get-VBOOrganization).username
    $OrgType = (Get-VBOOrganization).type
    $OrgService = (Get-VBOOrganization).BackupParts
    $OrgRegion = (Get-VBOOrganization).region
    if ((Get-VBOOrganization).Office365ExchangeConnectionSettings -ne $null)
    {
        $OrgAuth = (Get-VBOOrganization).Office365ExchangeConnectionSettings.AuthenticationType
    }
    else
    {
        $OrgAuth = (Get-VBOOrganization).Office365SharePointConnectionSettings.AuthenticationType
    }
    if ($OrgAuth -eq "Basic")
    {
        $AuxAccount = (Get-VBOOrganization).backupaccounts.count

    }
    else
    {
        $AuxAccount = (Get-VBOOrganization).backupapplications.count

    }

    [PScustomObject]@{
        Name = $OrgName
        Account = $OrgAccount
        Type = $OrgType
        Service = $OrgService
        Region = $OrgRegion
        Authentication = $OrgAuth
        AuxAccount = $AuxAccount
    }
}

 <#
 .Synopsis
    Get configuration about backup job configuration
 .DESCRIPTION
    Get job name, type, included object, excluded object, repository, proxy, schedule, active or disabled state
 .EXAMPLE 
    Get-DCVBMBackupJob
 #>
function Get-DCVBMBackupJob
{
    Write-host "$(get-date -Format HH:mm) - VB365 Backup Jobs"

    foreach ($obj in Get-VBOJob)
    {
        $JobName = $obj.name
        $JobType = $obj.JobBackupType
        $JobIncludedObj = $obj.SelectedItems -join ","
        $JobExcludedObj = $obj.ExcludedItems -join ","
        $JobRepository = $obj.Repository
        #Get proxy name from associated proxy ID repository
        $JobProxy = (Get-VBOProxy -id (Get-VBORepository -name $obj.Repository).proxyid).Hostname
        $JobSchedule = "<N/A>"
        if ($obj.schedulepolicy.EnableSchedule -and $obj.schedulepolicy.type -eq "daily")
        {
             $JobSchedule = [string]$obj.SchedulePolicy.DailyTime + " " + $obj.SchedulePolicy.DailyType
        }
        if ($obj.schedulepolicy.EnableSchedule -and $obj.schedulepolicy.type -eq "Periodically")
        {
             $JobSchedule = $obj.SchedulePolicy.PeriodicallyEvery
        }
        $JobEnabled = $obj.IsEnabled


        [PScustomObject]@{
            Name = $JobName
            Type = $JobType
            InclObject = $JobIncludedObj
            ExclObject = $JobExcludedObj
            Repository = $JobRepository
            Proxy = $JobProxy
            Schedule = $JobSchedule
            Enabled = $JobEnabled
        }
    }
}

<#
 .Synopsis
    Get configuration about backup copy job configuration
 .DESCRIPTION
    Get job name, repository, backupjob linked, schedule, active or disabled state
 .EXAMPLE 
    Get-DCVBMBackupCopyJob
 #>
function Get-DCVBMBackupCopyJob
{
    Write-host "$(get-date -Format HH:mm) - VB365 Backup copy Jobs"

    foreach ($obj in Get-VBOCopyJob)
    {
        $JobName = $obj.name
        $JobRepository = $obj.Repository
        $JobBackupLinked = $obj.BackupJob
        if ($obj.schedulepolicy.type -eq "daily")
        {
             $JobSchedule = [string]$obj.SchedulePolicy.DailyTime + " " + $obj.SchedulePolicy.DailyType
        }
        if ($obj.schedulepolicy.type -eq "Periodically")
        {
             $JobSchedule = $obj.SchedulePolicy.PeriodicallyEvery
        }
        else 
        {
            $JobSchedule = $obj.SchedulePolicy.Type
        }
        $JobEnabled = $obj.IsEnabled

        [PScustomObject]@{
            Name = $JobName
            Repository = $JobRepository
            BackupLinked = $JobBackupLinked
            Schedule = $JobSchedule
            Enabled = $JobEnabled
        }
    }
}


 <#
 .Synopsis
    Get configuration about proxy configuration
 .DESCRIPTION
    Get proxy name, port, thread number, throttling, internet proxy used or not, internet proxy port and account
 .EXAMPLE 
    Get-DCVBMProxy
 #>
function Get-DCVBMProxy
{

    Write-host "$(get-date -Format HH:mm) - VB365 Proxy"

    foreach ($obj in Get-VBOProxy)
    {
        $ProxyName = $obj.hostname
        $ProxyPort = $obj.port
        $ProxyThread = $obj.ThreadsNumber
        $ProxyThrottling = [string]$obj.ThrottlingValue + " " + $obj.ThrottlingUnit
        $ProxyIntHost = "<N/A>"
        $ProxyInternetPort = "<N/A>"
        $ProxyInternetAccount = "<N/A>"
        if ($obj.InternetProxy.UseInternetProxy)
        {
            $ProxyIntHost = $obj.InternetProxy.UseInternetProxy.Host
            $ProxyInternetPort = $obj.InternetProxy.UseInternetProxy.Port
            $ProxyInternetAccount = $obj.InternetProxy.UseInternetProxy.User
        }

        [PScustomObject]@{
            Name = $ProxyName
            Port = $ProxyPort
            Thread = $ProxyThread
            Throttling = $ProxyThrottling
            IntProxyHost = $ProxyIntHost
            IntProxyPort = $ProxyInternetPort
            IntProxyAccount = $ProxyInternetAccount
        }
    }
}

 <#
 .Synopsis
    Get configuration about repository configuration
 .DESCRIPTION
    Get repository name, proxy associated, path, retention type and value, repository object name and encryption
 .EXAMPLE 
    Get-DCVBMRepository
 #>
function Get-DCVBMRepository
{

    Write-host "$(get-date -Format HH:mm) - VB365 Repository"

    foreach ($obj in Get-VBORepository)
    {
        $RepositoryName = $obj.name
        $RepositoryProxy =  (Get-VBOProxy -id (Get-VBORepository -name $obj.name).proxyid).Hostname
        $RepositoryPath = $obj.Path
        #En attente forum https://forums.veeam.com/veeam-backup-for-microsoft-365-f47/powershell-host-repository-t81718.html
        $RepositoryRetention = [string]$obj.retentionperiod + " " + $obj.RetentionType
        $RepositoryObjectName = "<N/A>"
        if ($obj.ObjectStorageRepository -ne $null)
        {
            $RepositoryObjectName = $obj.ObjectStorageRepository.Name
        }
        $encryption = $obj.EnableObjectStorageEncryption
        [PScustomObject]@{
            Name = $RepositoryName
            Proxy = $RepositoryProxy
            Path = $RepositoryPath
            ObjectRepository = $RepositoryObjectName
            Retention = $RepositoryRetention
            Encryption = $encryption
        }
    }
}

 <#
 .Synopsis
    Get configuration about object repository configuration
 .DESCRIPTION
    Get repository name, folder, type, size limit and if it's long term achive
 .EXAMPLE 
    Get-DCVBMRepository
 #>
function Get-DCVBMObjectRepository
{

    Write-host "$(get-date -Format HH:mm) - VBM365 Object Repository"

    foreach ($obj in Get-VBOObjectStorageRepository)
    {
        $RepositoryName = $obj.name
        $RepositoryFolder =  $obj.Folder
        $RepositoryType = $obj.Type
        $RepositorySizeLimit = "<N/A>"
        if ($obj.EnableSizeLimit)
        {
            $RepositorySizeLimit = [String]$obj.UsedSpace + "/" + $obj.SizeLimit
        }
        $RepositoryArchive = $obj.IsLongTerm
        [PScustomObject]@{
            Name = $RepositoryName
            Folder = $RepositoryFolder
            Type = $RepositoryType
            SizeLimit = $RepositorySizeLimit
            LongTerm = $RepositoryArchive
        }
    }
}

 <#
 .Synopsis
    Get configuration about license
 .DESCRIPTION
    Get license type, expiration date, customer, contact, usage
 .EXAMPLE 
    Get-DCVBMLicense
 #>
function Get-DCVBMLicense
{

    Write-host "$(get-date -Format HH:mm) - VB365 License"

        $LicenseType = (Get-VBOLicense).type
        $LicenseExpiration =  (Get-VBOLicense).expirationdate.ToShortDateString()
        $LicenseTo = (Get-VBOLicense).LicensedTo
        $LicenseContact = (Get-VBOLicense).ContactPerson
        $LicenseUser = [string](Get-VBOLicense).usedNumber + "/" + (Get-VBOLicense).TotalNumber

        [PScustomObject]@{
            Type = $LicenseType
            Expiration = $LicenseExpiration
            To = $LicenseTo
            Contact = $LicenseContact
            Number = $LicenseUser
        }
}

 <#
 .Synopsis
    Get configuration about restore operator configuration
 .DESCRIPTION
    Get role name, organization, operator, associated object, excluded object
 .EXAMPLE 
    Get-DCVBMRestoreOperator
 #>
function Get-DCVBMRestoreOperator
{
    Write-host "$(get-date -Format HH:mm) - VB365 Restore Operator"

    foreach ($obj in Get-VBORbacRole)
    {
        $RoleName = $obj.name
        $OrganizationName =  (Get-VBOOrganization -Id ($obj.OrganizationId)).Name
        $OperatorName = $obj.operators.DisplayName -join ","
        $IncludedObject = "Organization"
        if ($obj.RoleType -ne "EntireOrganization")
        {
            $IncludedObject = $obj.SelectedItems.DisplayName -join ","
        }
        $ExcludedObject = "<N/A>"
         if ($obj.ExcludedItems -ne $null)
         {
            $ExcludedObject = $obj.ExcludedItems.DisplayName -join ","
         }
        [PScustomObject]@{
            Role = $RoleName
            Organization = $OrganizationName
            Operator = $OperatorName
            IncludedObject = $IncludedObject
            ExcludedObject = $ExcludedObject
        }
    }
}

 <#
 .Synopsis
    Get configuration about RestAPI configuration
 .DESCRIPTION
    Get state, token life time, port, certificate thumbprint friendly name and expiration date
 .EXAMPLE 
    Get-DCVBMRestAPI
 #>
function Get-DCVBMRestAPI
{
    Write-host "$(get-date -Format HH:mm) - VB365 REST API"

        $Enabled = (Get-VBORestAPISettings).IsServiceEnabled
        $TokenTime =  (Get-VBORestAPISettings).AuthTokenLifeTime
        $Port = (Get-VBORestAPISettings).HTTPSPort
        $CertThumbprint = (Get-VBORestAPISettings).CertificateThumbprint
        $CertFriendlyName = (Get-VBORestAPISettings).CertificateFriendlyName
        $CertExpiration = (Get-VBORestAPISettings).CertificateExpirationDate.ToShortDateString()

        [PScustomObject]@{
            Enabled = $Enabled
            TokenTime = $TokenTime
            Port = $Port
            CertThumbprint = $CertThumbprint
            CertFriendlyName = $CertFriendlyName
            CertExpiration = $CertExpiration
        }
}


 <#
 .Synopsis
    Get configuration about Restore portal configuration
 .DESCRIPTION
    Get state, application ID, certificate thumbprint friendly name and expiration date
 .EXAMPLE 
    Get-DCVBMRestorePortal
 #>
function Get-DCVBMRestorePortal
{

    Write-host "$(get-date -Format HH:mm) - VB365 Restore portal"

    $Enabled = (Get-VBORestorePortalSettings).IsServiceEnabled
    $ApplicationID =  (Get-VBORestorePortalSettings).ApplicationId.Guid
    $CertThumbprint = (Get-VBORestorePortalSettings).CertificateThumbprint
    $CertFriendlyName = (Get-VBORestorePortalSettings).CertificateFriendlyName
    $CertExpiration = (Get-VBORestorePortalSettings).CertificateExpirationDate.ToShortDateString()

    [PScustomObject]@{
        Enabled = $Enabled
        ApplicationID = $ApplicationID
        CertThumbprint = $CertThumbprint
        CertFriendlyName = $CertFriendlyName
        CertExpiration = $CertExpiration
    }
}

 <#
 .Synopsis
    Get configuration about operator Authentication portal configuration
 .DESCRIPTION
    Get state, certificate thumbprint friendly name and expiration date
 .EXAMPLE 
    Get-DCVBMOperatorAuthentication
 #>
function Get-DCVBMOperatorAuthentication
{

    Write-host "$(get-date -Format HH:mm) - VB365 Authentication"

    $Enabled = (Get-VBOOperatorAuthenticationSettings).AuthenticationEnabled
    $CertThumbprint = (Get-VBORestAPISettings).CertificateThumbprint
    $CertFriendlyName = (Get-VBOOperatorAuthenticationSettings).CertificateFriendlyName
    $CertExpiration = (Get-VBOOperatorAuthenticationSettings).CertificateExpirationDate

    [PScustomObject]@{
        Enabled = $Enabled
        CertThumbprint = $CertThumbprint
        CertFriendlyName = $CertFriendlyName
        CertExpiration = $CertExpiration
    }
}

 <#
 .Synopsis
    Get configuration about internet proxy
 .DESCRIPTION
    Get state, host, port and account
 .EXAMPLE 
    Get-DCVBMInternetProxy
 #>
function Get-DCVBMInternetProxy
{

    Write-host "$(get-date -Format HH:mm) - VB365 Internet Proxy"

    $IntProxyEnabled = (Get-VBOInternetProxySettings).UseInternetProxy
    $IntProxyHost = "<N/A>"
    $IntProxyPort = "<N/A>"
    $IntProxyUser = "<N/A>"
    if ((Get-VBOInternetProxySettings).UseInternetProxy)
    {
        $IntProxyHost = (Get-VBOInternetProxySettings).Host
        $IntProxyPort = (Get-VBOInternetProxySettings).Port
        $IntProxyUser = (Get-VBOInternetProxySettings).User
    }

    [PScustomObject]@{
        Enabled = $IntProxyEnabled
        Host = $IntProxyHost
        Port = $IntProxyPort
        Account = $IntProxyUser
    }
}

 <#
 .Synopsis
    Get configuration about SMTP
 .DESCRIPTION
    Get state, server, port, ssl, account
 .EXAMPLE 
    Get-DCVBMSMTP
 #>
function Get-DCVBMSMTP
{

    Write-host "$(get-date -Format HH:mm) - VB365 SMTP configuration"

    $SMTPEnabled = (Get-VBOEmailSettings).EnableNotification
    $SMTPServer = "<N/A>"
    $SMTPPort = "<N/A>"
    $SMTPSSL = "<N/A>"
    $SMTPAccount = "<N/A>"
    if ((Get-VBOEmailSettings).EnableNotification)
    {
        $SMTPServer = (Get-VBOEmailSettings).SMTPServer
        $SMTPPort = (Get-VBOEmailSettings).Port
        $SMTPSSL = (Get-VBOEmailSettings).UseSSL
        if ((Get-VBOEmailSettings).UseAuthentication)
        {
            $SMTPAccount = (Get-VBOEmailSettings).Username
        }
    }

    [PScustomObject]@{
        Enabled = $SMTPEnabled
        Server = $SMTPServer
        Port = $SMTPPort
        SSL = $SMTPSSL
        Account = $SMTPAccount
    }
}


 <#
 .Synopsis
    Get configuration about Notifications
 .DESCRIPTION
    Get state, sender, receiver, notification on success, warning and failure, send only last retry notification
 .EXAMPLE 
    Get-DCVBMNotification
 #>
function Get-DCVBMNotification
{

    Write-host "$(get-date -Format HH:mm) - VB365 Notifications"

    $NotificationEnabled = (Get-VBOEmailSettings).EnableNotification
    $NotificationSender = "<N/A>"
    $NotificationReceiver = "<N/A>"
    $NotificationSuccess = "<N/A>"
    $NotificationWarning = "<N/A>"
    $NotificationFailure = "<N/A>"
    $LastRetryNotificationOnly = "<N/A>"
    if ((Get-VBOEmailSettings).EnableNotification)
    {
        $NotificationSender = (Get-VBOEmailSettings).From -join ","
        $NotificationReceiver = (Get-VBOEmailSettings).To -join ","
        $NotificationSuccess = (Get-VBOEmailSettings).NotifyOnSuccess
        $NotificationWarning = (Get-VBOEmailSettings).NotifyOnWarning
        $NotificationFailure = (Get-VBOEmailSettings).NotifyOnFailure
        $LastRetryNotificationOnly = (Get-VBOEmailSettings).SupressUntilLastRetry
    }

    [PScustomObject]@{
        Enabled = $NotificationEnabled
        Sender = $NotificationSender
        Receiver = $NotificationReceiver
        OnSuccess = $NotificationSuccess
        OnWarning = $NotificationWarning
        OnFailure = $NotificationFailure
        OnlyLastRetry = $LastRetryNotificationOnly
    }
}


 <#
 .Synopsis
    Create array for HTML report
 .DESCRIPTION
    Create array with title and precontent
 .EXAMPLE 
    CreateArray -title "my Title" -var $MyData -PreContent $MyPrecontent
 #>
Function CreateArray ($Title,$Var,$PreContent)
{
    if ($Title)
    {
        "<h3>$Title</h3>"
    }
    if ($PreContent)
    {
        $Var | ConvertTo-Html -Fragment -PreContent $PreContent
    }
    else
    {
        $Var | ConvertTo-Html -Fragment
    }
}


<#
.Synopsis
   Generate HTML report
.DESCRIPTION
   Use all variable to build html report with CSS style 
.EXAMPLE
   Get-HTMLReport -Path "c:\temp\report.html"
#>

function Get-HTMLReport
{
    Write-Host "Building HTML"

    #region HTML
    @"
    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head>
    <title>$HTMLTitle</title>
    $HTMLCSS
    </head>
    <body>
    <br><br><br><br>

    <h1>VEEAM Backup for Microsoft 365 Report</h1>

    $(CreateArray -title "Summary" -var $DCVBMSummary) 
    $(CreateArray -title "License" -var $DCVBMLicense)
    $(CreateArray -title "SMTP" -var $DCVBMSMTP)
    $(CreateArray -title "Notifications" -var $DCVBMNotification)
    $(CreateArray -title "Internet Proxy" -var $DCVBMInternetProxy)
    
    $(CreateArray -title "REST API" -var $DCVBMRestAPI)
    $(CreateArray -title "Restore portal" -var $DCVBMRestorePortal)
    $(CreateArray -title "Get-DCVBMOperatorAuthentication" -var $DCVBMOperatorAuthentication )

    $(CreateArray -title "Repositories" -var $DCVBMRepository)
    $(CreateArray -title "Object Repositories" -var $DCVBMObjectRepository)
        
    $(CreateArray -title "Proxies" -var $DCVBMProxy)

    $(CreateArray -title "Organizations" -var $DCVBMOrganization)
    $(CreateArray -title "Backup jobs" -var $DCVBMBackupJob)
    $(CreateArray -title "Backup copy jobs" -var $DCVBMBackupCopyJob)

    $(CreateArray -title "Restore operators" -var $DCVBMRestoreOperator )
    
    </body>
"@ | Out-File -Encoding utf8 $htmlReportPath

    Invoke-Item $htmlReportPath 
}
#endregion

#Write here all function that need to be displayed in all reports types

$DCVBMSummary = Get-DCVBMSummary
$DCVBMOrganization = Get-DCVBMOrganization
$DCVBMBackupJob = Get-DCVBMBackupJob
$DCVBMProxy = Get-DCVBMProxy
$DCVBMRepository = Get-DCVBMRepository
$DCVBMLicense = Get-DCVBMLicense
$DCVBMRestoreOperator = Get-DCVBMRestoreOperator
$DCVBMRestAPI = Get-DCVBMRestAPI
$DCVBMRestorePortal = Get-DCVBMRestorePortal
$DCVBMOperatorAuthentication = Get-DCVBMOperatorAuthentication
$DCVBMInternetProxy = Get-DCVBMInternetProxy
$DCVBMSMTP = Get-DCVBMSMTP
$DCVBMNotification = Get-DCVBMNotification
$DCVBMObjectRepository = Get-DCVBMObjectRepository
$DCVBMBackupCopyJob = Get-DCVBMBackupCopyJob

Get-HTMLReport 


Disconnect-VBOServer
