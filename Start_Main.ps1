<#
.SYNOPSIS
    . Script/Automation that grabs Microsoft Graph API Reporting data and dumps it to Azure SQL database. 
    Practical use would be: 
        . Some folks are not comfortable with API interactions. It's always easy to just get the data from SQL Database
        . Some reporting reporting dashboars don't have API query functionalities (or at least super complex to configure).
            . Most of reporting dashboards e.g. Grafana support natively SQL as a data source.
.DESCRIPTION
    Long description
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    . Prerequisites: 
        . Module: Logging - To install "Install-Module Logging -verbose -Force"
        . Module: SQLSever - To install "Install-Module SqlServer -verbose -Force"
        . Azure App:
            . Application ID and Secret key
            . Permissions to Graph API Reporting

    . Use the Create-Database script file to create the:
        . Resource Group
        . SQL Server
        . Database
        . User
        . Firewall exceptions
        
        * You can skip this part of you've created your DB already. 
    
    . Comment/Uncomment the Create-Table function at the bottom of the script when
        Creating or after creating the tables (respectively). You only need this one time. 

    . I have tested this on Windows 10 and Windows Server 2012 R2

    . Passwords and Secrets - prior to running the script, make sure to run the Create-Configuration
        first. The script will ask for all the necessary information and encrypt the passwords and
        secrets to a .json file. The .json file will picked by the main script for use. 

        *The encripted passwords can only be decrypted by the same user account who encrypted it
            when using the script on a different account, re-run the create-Configuration script.
        
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 


#region Logging Config
Import-module Logging # Please make sure to install this module
Set-LoggingDefaultLevel -Level INFO
Add-LoggingTarget -Name Console
Add-LoggingTarget -Name File -Configuration @{Path = "$psScriptRoot\$(Get-Date -Format "MMddyy") - Log.log" }
#endregion logging config

#region variables
#For this portion to work, make sure to run Create-Configuration.
$configuration = Get-Content .\configuration.json | ConvertFrom-Json

$Script:AzureAppIDandSecret = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($configuration.AzureAppID, ($configuration.AzureAppKey | ConvertTo-SecureString))
$Script:SQLCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($configuration.sqlCredentialUserName, ($configuration.sqlCredentialPassword | ConvertTo-SecureString))

$script:DBName = $configuration.DBName
$script:dbServer = $configuration.DbServer
$Script:tenantID = $configuration.TenantID

#endregion Variables

#region Table Creation
Function Create-Table { 
    Function New-GrafanaSQLTable { 

        [CMDletbinding()]

        param(
            [String]$dbServer,
            [String]$dbName,
            [ValidateNotNull()]
            [System.Management.Automation.PSCredential]
            [System.Management.Automation.Credential()]
            $Credential = [System.Management.Automation.PSCredential]::Empty,

            [Switch]$All,

            [Switch]$getTeamsDeviceUsageUserDetail,
            [Switch]$getTeamsDeviceUsageUserCounts,
            [Switch]$getTeamsDeviceUsageDistributionUserCounts,

            [Switch]$getTeamsUserActivityUserDetail,
            [Switch]$getTeamsUserActivityCounts,
            [Switch]$getTeamsUserActivityUserCounts,

            [Switch]$getEmailActivityUserDetail,
            [Switch]$getEmailActivityCounts,
            [Switch]$getEmailActivityUserCounts,

            [Switch]$getEmailAppUsageUserDetail,
            [Switch]$getEmailAppUsageAppsUserCounts,
            [Switch]$getEmailAppUsageUserCounts,
            [Switch]$getEmailAppUsageVersionsUserCounts,

            [Switch]$getMailboxUsageDetail,
            [Switch]$getMailboxUsageMailboxCounts,
            [Switch]$getMailboxUsageQuotaStatusMailboxCounts,
            [Switch]$getMailboxUsageStorage,

            [Switch]$getOffice365ActivationsUserDetail,
            [Switch]$getOffice365ActivationCounts,
            [Switch]$getOffice365ActivationsUserCounts,

            [Switch]$getOffice365ActiveUserDetail,
            [Switch]$getOffice365ActiveUserCounts,
            [Switch]$getOffice365ServicesUserCounts,

            [Switch]$getOffice365GroupsActivityDetail,
            [Switch]$getOffice365GroupsActivityCounts,
            [Switch]$getOffice365GroupsActivityGroupCounts,
            [Switch]$getOffice365GroupsActivityStorage,
            [Switch]$getOffice365GroupsActivityFileCounts
       
        ) 
        #Logging
        $SN = $MyInvocation.MyCommand.Name; Function WL($LE) { $LN = (Get-Date -Format "MMddyy:HHmmss") + " - $LE"; $LN | Out-File -FilePath "$PSScriptRoot\$SN-log.txt" -Append -NoClobber -Encoding "Default"; $LN }
        $script:ReportData = @()
    
        $SQLParams = @{ 
            ServerInstance = $SQLServer
            Database       = $DBName
            Credential     = $Credential
        }

        try {
            $ExistingTables = Invoke-Sqlcmd @SQLParams -Query "Select Table_Name from Information_Schema.Tables" -AbortOnError
 
        }
        Catch { 
            $Error[0].exception.Message
            WL "Error on fetching existing tables"
        
        }
     
        function Invoke-SQLQuery ([String]$QueryString, $TableName) {

            Try {
                WL "Invoking: $QueryString" 
                Invoke-Sqlcmd @SQLParams -Query $QueryString -AbortOnError -Verbose
                if ($?) {
                    $script:ReportData += [PSCustomObject]@{
                        Table_Name = $TableName
                        Result     = "Added Successfully"
                    }
                    WL "Invoking Success"        
                }
            }
            Catch {
                WL $Error[0].Exception.Message
                $script:ReportData += [PSCustomObject]@{
                    Table_Name = $TableName
                    Result     = "Error on Adding"
                }     
            }
        }

        Function Send-CreateTable ($Query, $TableName) { 

            if (-not ($ExistingTables.Table_Name.Contains($TableName))) {

                WL "Creating Table for $TableName"

                Invoke-SQLQuery -QueryString $Query -TableName $TableName           
            }

            Else { 
                WL "Table $TableName is already existing"
            
                $script:ReportData += [PSCustomObject]@{
                    Table_Name = $TableName
                    Result     = "Existing-Skipped"
                }
            }
        }

        # --> START OF TEAMS DEVICE USAGE <-- #
        # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#
        if ($getTeamsDeviceUsageUserDetail -or $all) {

            # Report contains all users - Span of 180 days
            $TableName = "getTeamsDeviceUsageUserDetail"
            $Query = @"
            CREATE TABLE $TableName (
                [User Principal Name] VARCHAR(50) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Last Activity Date] DATETIME,
                [Is Deleted] VARCHAR(10),
                [Deleted Date] VARCHAR(10),
                [Used Web] VARCHAR(10),
                [Used Windows Phone] VARCHAR(10),
                [Used iOS] VARCHAR(10),
                [Used Mac] VARCHAR(10),
                [Used Android Phone] VARCHAR(10),
                [Used Windows] VARCHAR(10),
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getTeamsDeviceUsageUserCounts -or $all) {

            $TableName = "getTeamsDeviceUsageUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Web] INT,
                [Windows Phone] INT,
                [Android Phone] INT,
                [iOS] INT,
                [Mac] INT,
                [Windows] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getTeamsDeviceUsageDistributionUserCounts -or $all) {
            # Single row report - Span of 180 days
            $TableName = "getTeamsDeviceUsageDistributionUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Refresh Date] DATETIME,
                [Web] INT,
                [Windows Phone] INT,
                [Android Phone] INT,
                [iOS] INT,
                [Mac] INT,
                [Windows] INT,
                [Report Period] INT            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }
    
        # --> END OF TEAMS TEAMS DEVICE USAGE <-- #
        # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#
        # -->    START OF OF TEAMS REPORTS    <-- #
        if ($getTeamsUserActivityUserDetail -or $all) {

            $TableName = "getTeamsUserActivityUserDetail"
            $Query = @"
            CREATE TABLE $TableName (

                [User Principal Name] VARCHAR(50) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Last Activity Date] DATETIME,
                [Is Deleted] VARCHAR(50),
                [Deleted Date] DATETIME,
                [Assigned Products] VARCHAR(500),
                [Team Chat Message Count] INT,
                [Private Chat Message Count] INT,
                [Call Count] INT,
                [Meeting Count] INT,
                [Has Other Action] VARCHAR(50),
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }
    
        if ($getTeamsUserActivityCounts -or $all) {

            $TableName = "getTeamsUserActivityCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Team Chat Messages] INT,
                [Private Chat Messages] INT,
                [Calls] INT,
                [Meetings] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getTeamsUserActivityUserCounts -or $all) {

            $TableName = "getTeamsUserActivityUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Team Chat Messages] INT,
                [Private Chat Messages] INT,
                [Calls] INT,
                [Meetings] INT,
                [Other Actions] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }
    
        # -->   END OF TEAMS  USER ACTIVITY   <-- #
        # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#
        # -->   START OF OUTLOOK ACTIVITY     <-- #

        if ($getEmailActivityUserDetail -or $all) {
            # An all users Report
            $TableName = "getEmailActivityUserDetail"
            $Query = @"
            CREATE TABLE $TableName (
                [User Principal Name] VARCHAR(50) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Display Name] VARCHAR(250),
                [Is Deleted] VARCHAR(10),
                [Deleted Date] DATETIME,
                [Last Activity Date] DATETIME,
                [Send Count] INT,
                [Receive Count] INT,
                [Read Count] INT,
                [Assigned Products] VARCHAR(500), 
                [Report Period] INT          
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }
    
        if ($getEmailActivityCounts -or $all) {

            $TableName = "getEmailActivityCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Send] INT,
                [Receive] INT,
                [Read] INT,
                [Report Period] INT,
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getEmailActivityUserCounts -or $all) {

            $TableName = "getEmailActivityUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Send] INT,
                [Receive] INT,
                [Read] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        # -->  END OF TEAMS OUTLOOK ACTIVITY  <-- #
        # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#
        # -->  START OF OF OUTLOOK APP USAGE  <-- #
        if ($getEmailAppUsageUserDetail -or $all) {
            # All user report
            $TableName = "getEmailAppUsageUserDetail"
            $Query = @"
            CREATE TABLE $TableName (
                [User Principal Name] VARCHAR(50) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Display Name] VARCHAR(250),
                [Is Deleted] VARCHAR(10),
                [Deleted Date] DATETIME,
                [Last Activity Date] DATETIME,
                [Mail For Mac] VARCHAR (50),
                [Outlook For Mac] VARCHAR (50),
                [Outlook For Windows] VARCHAR (50),
                [Outlook For Mobile] VARCHAR (50),
                [Other For Mobile] VARCHAR (50),
                [Outlook For Web] VARCHAR (50),
                [POP3 App] VARCHAR (50),
                [IMAP4 App] VARCHAR (50),
                [SMTP App] VARCHAR (50),
                [Report Period] INT  
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getEmailAppUsageAppsUserCounts -or $all) {
            # One line report
            $TableName = "getEmailAppUsageAppsUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Refresh Date] DATE NOT NULL,
                [Mail For Mac] INT,
                [Outlook For Mac] INT,
                [Outlook For Windows] INT,
                [Outlook For Mobile] INT,
                [Other For Mobile] INT,
                [Outlook For Web] INT,
                [POP3 App] INT,
                [IMAP4 App] INT,
                [SMTP App] INT,
                [Report Period]  INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getEmailAppUsageUserCounts -or $all) {

            $TableName = "getEmailAppUsageUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME,
                [Report Refresh Date] DATETIME,
                [Mail For Mac] INT,
                [Outlook For Mac] INT,
                [Outlook For Windows] INT,
                [Outlook For Mobile] INT,
                [Other For Mobile] INT,
                [Outlook For Web] INT,
                [POP3 App] INT,
                [IMAP4 App] INT,
                [SMTP App] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getEmailAppUsageVersionsUserCounts -or $all) {

            $TableName = "getEmailAppUsageVersionsUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Refresh Date] DATE NOT NULL,
                [Outlook 2016] INT,
                [Outlook 2013] INT,
                [Outlook 2010] INT,
                [Outlook 2007] INT,
                [Undetermined] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        # -->      END OF OUTLOOK APP USAGE        <-- #
        # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-#
        # -->  START OF OF OUTLOOK MAILBOX USAGE   <-- #
    
        if ($getMailboxUsageDetail -or $all) {

            $TableName = "getMailboxUsageDetail"
            $Query = @"
            
            CREATE TABLE $TableName (
                [User Principal Name] VARCHAR(50) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Display Name] VARCHAR(100),
                [Is Deleted] VARCHAR(10),
                [Deleted Date] DATETIME,
                [Created Date] DATETIME,
                [Last Activity Date] DATETIME,
                [Item Count] INT,
                [Storage Used (Byte)] BIGINT,
                [Issue Warning Quota (Byte)] BIGINT,
                [Prohibit Send Quota (Byte)] BIGINT,
                [Prohibit Send/Receive Quota (Byte)] BIGINT,
                [Deleted Item Count] INT,
                [Deleted Item Size (Byte)] BIGINT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getMailboxUsageMailboxCounts -or $all) {

            $TableName = "getMailboxUsageMailboxCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Total] INT,
                [Active] INT,
                [Report Period] INT   
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getMailboxUsageQuotaStatusMailboxCounts -or $all) {

            $TableName = "getMailboxUsageQuotaStatusMailboxCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Under Limit] INT,
                [Warning Issued] INT,
                [Send Prohibited] INT,
                [Send/Receive Prohibited] INT,
                [Indeterminate] INT,
                [Report Period] INT,     
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getMailboxUsageStorage -or $all) {

            $TableName = "getMailboxUsageStorage"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Storage Used (Byte)] BigInt,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        # -->   END OF OUTLOOK MAILBOX USAGE    <-- #
        # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-==--=-=-#
        # -->  START OF OFFICE 365 ACTIVATIONS  <-- #

        if ($getOffice365ActivationsUserDetail -or $all) {

            $TableName = "getOffice365ActivationsUserDetail"
            $Query = @"
            CREATE TABLE $TableName (
                [User Principal Name] VARCHAR(50) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Display Name] VARCHAR(50),
                [Product Type] VARCHAR (100),
                [Last Activated Date] DATETIME,
                [Windows] INT,
                [Mac] INT,
                [Windows 10 Mobile] INT,
                [iOS] INT,
                [Android] INT,
                [Activated On Shared Computer] VARCHAR(10)                       
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOffice365ActivationCounts -or $all) {

            $TableName = "getOffice365ActivationCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Product Type] VARCHAR(100) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Windows] INT,
                [Mac] INT,
                [Android] INT,
                [iOS] INT,
                [Windows 10 Mobile] INT            
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOffice365ActivationsUserCounts -or $all) {

            $TableName = "getOffice365ActivationsUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Product Type] VARCHAR (100) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Assigned] INT,
                [Activated] INT,
                [Shared Computer Activation] VARCHAR(10)       
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        # -->   END OF OFFICE 365 ACTIVATIONS    <-- #
        # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-#
        # -->  START OF OFFICE 365 ACTIVE USERS  <-- #

        if ($getOffice365ActiveUserDetail -or $all) {

            $TableName = "getOffice365ActiveUserDetail"
            $Query = @"
            CREATE TABLE $TableName ( 
                [User Principal Name] VARCHAR(100) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Display Name] VARCHAR(250),
                [Is Deleted] VARCHAR(10),
                [Deleted Date] DATETIME,
                [Has Exchange License] VARCHAR(10),
                [Has OneDrive License] VARCHAR(10),
                [Has SharePoint License] VARCHAR(10),
                [Has Skype For Business License] VARCHAR(10),
                [Has Yammer License] VARCHAR(10),
                [Has Teams License] VARCHAR(10),
                [Exchange Last Activity Date] VARCHAR(10),
                [OneDrive Last Activity Date] DATETIME,
                [SharePoint Last Activity Date] DATETIME,
                [Skype For Business Last Activity Date] DATETIME,
                [Yammer Last Activity Date] DATETIME,
                [Teams Last Activity Date] DATETIME,
                [Exchange License Assign Date] DATETIME,
                [OneDrive License Assign Date] DATETIME,
                [SharePoint License Assign Date] DATETIME,
                [Skype For Business License Assign Date] DATETIME,
                [Yammer License Assign Date] DATETIME,
                [Teams License Assign Date] DATETIME,
                [Assigned Products] VARCHAR(100)      
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOffice365ActiveUserCounts -or $all) {

            $TableName = "getOffice365ActiveUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Office 365] INT,
                [Exchange] INT,
                [OneDrive] INT,
                [SharePoint] INT,
                [Skype For Business] INT,
                [Yammer] INT,
                [Teams] INT,
                [Report Period] INT       
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOffice365ServicesUserCounts -or $all) {

            $TableName = "getOffice365ServicesUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                    [Report Refresh Date] DATETIME,
                    [Exchange Active] INT,
                    [Exchange Inactive] INT,
                    [OneDrive Active] INT,
                    [OneDrive Inactive] INT,
                    [SharePoint Active] INT,
                    [SharePoint Inactive] INT,
                    [Skype For Business Active] INT,
                    [Skype For Business Inactive] INT,
                    [Yammer Active] INT,
                    [Yammer Inactive] INT,
                    [Teams Active] INT,
                    [Teams Inactive] INT,
                    [Office 365 Active] INT,
                    [Office 365 Inactive] INT,
                    [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        # -->   END OF  OFFICE 365 ACTIVE USERS <-- #
        # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#
        # -->   START OF O365 GROUPS ACTIVITY   <-- #

        if ($getOffice365GroupsActivityDetail -or $all) {

            $TableName = "getOffice365GroupsActivityDetail"
            $Query = @"
            CREATE TABLE $TableName (
                [Group Id] VARCHAR(100) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Group Display Name] VARCHAR(100),
                [Is Deleted] VARCHAR(10),
                [Owner Principal Name] VARCHAR(100),
                [Last Activity Date] DATETIME,
                [Group Type] VARCHAR(20),
                [Member Count] INT,
                [External Member Count] INT,
                [Exchange Received Email Count] INT,
                [SharePoint Active File Count] INT,
                [Yammer Posted Message Count] INT,
                [Yammer Read Message Count] INT,
                [Yammer Liked Message Count] INT,
                [Exchange Mailbox Total Item Count] INT,
                [Exchange Mailbox Storage Used (Byte)] BIGINT,
                [SharePoint Total File Count] INT,
                [SharePoint Site Storage Used (Byte)] BIGINT,
                [Report Period] INT    
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOffice365GroupsActivityCounts -or $all) {

            $TableName = "getOffice365GroupsActivityCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME,
                [Report Refresh Date] DATETIME,
                [Exchange Emails Received] INT,
                [Yammer Messages Posted] INT,
                [Yammer Messages Read] INT,
                [Yammer Messages Liked] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOffice365GroupsActivityGroupCounts -or $all) {

            $TableName = "getOffice365GroupsActivityGroupCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME,
                [Report Refresh Date] DATETIME,
                [Total] INT,
                [Active] INT,
                [Report Period] INT        
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOffice365GroupsActivityStorage -or $all) {

            $TableName = "getOffice365GroupsActivityStorage"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME,
                [Report Refresh Date] DATETIME,
                [Mailbox Storage Used (Byte)] BIGINT,
                [Site Storage Used (Byte)] BIGINT,
                [Report Period] INT          
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOffice365GroupsActivityFileCounts -or $all) {

            $TableName = "getOffice365GroupsActivityFileCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME,
                [Report Refresh Date] DATETIME,
                [Total] INT,
                [Active] INT,
                [Report Period] INT      
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOneDriveActivityUserDetail -or $all) {

            $TableName = "getOneDriveActivityUserDetail"
            $Query = @"
            CREATE TABLE $TableName (
                [User Principal Name] VARCHAR(100) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Is Deleted] VARCHAR(10),
                [Deleted Date] DATETIME,
                [Last Activity Date] DATETIME,
                [Viewed Or Edited File Count] INT,
                [Synced File Count] INT,
                [Shared Internally File Count] INT,
                [Shared Externally File Count] INT,
                [Assigned Products] VARCHAR(250),
                [Report Period] INT

            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOneDriveActivityUserCounts -or $all) {

            $TableName = "getOneDriveActivityUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Viewed Or Edited] INT,
                [Synced] INT,
                [Shared Internally] INT,
                [Shared Externally]INT, 
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }
    
        if ($getOneDriveActivityFileCounts -or $all) {

            $TableName = "getOneDriveActivityFileCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Viewed Or Edited] INT,
                [Synced] INT,
                [Shared Internally] INT,
                [Shared Externally] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOneDriveUsageAccountDetail -or $all) {

            $TableName = "getOneDriveUsageAccountDetail"
            $Query = @"
            CREATE TABLE $TableName (
                [Owner Principal Name] VARCHAR(150) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Site URL] VARCHAR(500),
                [Owner Display Name] VARCHAR(250),
                [Is Deleted] VARCHAR(10),
                [Last Activity Date] DATETIME,
                [File Count] INT,
                [Active File Count] INT,
                [Storage Used (Byte)] BIGINT,
                [Storage Allocated (Byte)] BIGINT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOneDriveUsageAccountCounts -or $all) {

            $TableName = "getOneDriveUsageAccountCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Site Type] VARCHAR(50),
                [Total] INT,
                [Active] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getOneDriveUsageFileCounts -or $all) {

            $TableName = "getOneDriveUsageFileCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,    
                [Report Refresh Date] DATETIME,
                [Site Type] VARCHAR(50),
                [Total] INT,
                [Active] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }
    
        if ($getOneDriveUsageStorage -or $all) {

            $TableName = "getOneDriveUsageStorage"
            $Query = @"
                CREATE TABLE $TableName (
                    [Report Date] DATETIME NOT NULL,
                    [Report Refresh Date] DATETIME,
                    [Site Type] VARCHAR(50),
                    [Storage Used (Byte)] BIGINT,
                    [Report Period] INT
                )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSharePointActivityUserDetail -or $all) {

            $TableName = "getSharePointActivityUserDetail"
            $Query = @"
            CREATE TABLE $TableName (
                [User Principal Name] VARCHAR(150) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Is Deleted] VARCHAR(10),
                [Deleted Date] DATETIME,
                [Last Activity Date] DATETIME,
                [Viewed Or Edited File Count] INT,
                [Synced File Count] INT,
                [Shared Internally File Count] INT,
                [Shared Externally File Count] INT,
                [Visited Page Count] INT,
                [Assigned Products] VARCHAR(250),
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSharePointActivityFileCounts -or $all) {

            $TableName = "getSharePointActivityFileCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Viewed Or Edited] INT,
                [Synced] INT,
                [Shared Internally] INT,
                [Shared Externally] INT,
                [Report Period] INT,
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSharePointActivityUserCounts -or $all) {

            $TableName = "getSharePointActivityUserCounts"
            $Query = @"
            CREATE TABLE $TableName ( 
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Visited Page] INT,
                [Viewed Or Edited] INT,
                [Synced] INT,
                [Shared Internally] INT,
                [Shared Externally] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSharePointActivityPages -or $all) {

            $TableName = "getSharePointActivityPages"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Visited Page Count] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSharePointSiteUsageDetail -or $all) {

            $TableName = "getSharePointSiteUsageDetail"
            $Query = @"
            CREATE TABLE $TableName (
                [Owner Principal Name] VARCHAR(150) NOT NULL,
                [Report Refresh Date] DATETIME,
                [Site Id] VARCHAR(100),
                [Site URL] VARCHAR(250),
                [Owner Display Name] VARCHAR(150),
                [Is Deleted] VARCHAR(10),
                [Last Activity Date] DATETIME,
                [File Count] INT,
                [Active File Count] INT,
                [Page View Count] INT,
                [Visited Page Count] INT,
                [Storage Used (Byte)] BIGINT,
                [Storage Allocated (Byte)] BIGINT,
                [Root Web Template] VARCHAR(50),
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSharePointSiteUsageFileCounts -or $all) {

            $TableName = "getSharePointSiteUsageFileCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Site Type] VARCHAR(50),
                [Total] INT,
                [Active] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSharePointSiteUsageSiteCounts -or $all) {

            $TableName = "getSharePointSiteUsageSiteCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Site Type] VARCHAR(50),
                [Total] INT,
                [Active] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSharePointSiteUsageStorage -or $all) {

            $TableName = "getSharePointSiteUsageStorage"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Site Type] VARCHAR(50),
                [Storage Used (Byte)]BIGINT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSharePointSiteUsagePages -or $all) {

            $TableName = "getSharePointSiteUsagePages"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME NOT NULL,
                [Report Refresh Date] DATETIME,
                [Site Type] VARCHAR(50),
                [Page View Count] INT,
                [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSkypeForBusinessActivityUserDetail -or $all) {

            $TableName = "getSkypeForBusinessActivityUserDetail"
            $Query = @"
            CREATE TABLE $TableName (
                    [User Principal Name] VARCHAR(150) NOT NULL,           
                    [Report Refresh Date] DATETIME,
                    [Is Deleted] VARCHAR(10),
                    [Deleted Date] DATETIME,
                    [Last Activity Date] DATETIME,
                    [Total Peer-to-peer Session Count] INT,
                    [Total Organized Conference Count] INT,
                    [Total Participated Conference Count] INT,
                    [Peer-to-peer Last Activity Date] DATETIME,
                    [Organized Conference Last Activity Date] DATETIME,
                    [Participated Conference Last Activity Date] DATETIME,
                    [Peer-to-peer IM Count] INT,
                    [Peer-to-peer Audio Count] INT,
                    [Peer-to-peer Audio Minutes] INT,
                    [Peer-to-peer Video Count] INT,
                    [Peer-to-peer Video Minutes] INT,
                    [Peer-to-peer App Sharing Count] INT,
                    [Peer-to-peer File Transfer Count] INT,
                    [Organized Conference IM Count] INT,
                    [Organized Conference Audio/Video Count] INT,
                    [Organized Conference Audio/Video Minutes] INT,
                    [Organized Conference App Sharing Count] INT,
                    [Organized Conference Web Count] INT,
                    [Organized Conference Dial-in/out 3rd Party Count] INT,
                    [Organized Conference Dial-in/out Microsoft Count] INT,
                    [Organized Conference Dial-in Microsoft Minutes] INT,
                    [Organized Conference Dial-out Microsoft Minutes] INT,
                    [Participated Conference IM Count] INT,
                    [Participated Conference Audio/Video Count] INT,
                    [Participated Conference Audio/Video Minutes] INT,
                    [Participated Conference App Sharing Count] INT,
                    [Participated Conference Web Count] INT,
                    [Participated Conference Dial-in/out 3rd Party Count] INT,
                    [Assigned Products] VARCHAR(250),
                    [Report Period] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSkypeForBusinessActivityCounts -or $all) {

            $TableName = "getSkypeForBusinessActivityCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME,
                [Report Refresh Date] DATETIME,
                [Report Period] INT,
                [Peer-to-peer] INT,
                [Organized] INT,
                [Participated] INT
                        )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        if ($getSkypeForBusinessActivityUserCounts -or $all) {

            $TableName = "getSkypeForBusinessActivityUserCounts"
            $Query = @"
            CREATE TABLE $TableName (
                [Report Date] DATETIME,
                [Report Refresh Date] DATETIME,
                [Report Period] INT,
                [Peer-to-peer] INT,
                [Organized] INT,
                [Participated] INT
            )
"@
            Send-CreateTable -Query $Query -TableName $TableName
        }

        # -->   END OF O365 GROUPS ACTIVITY     <-- #
        # --=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-#
        # -->   START OF O365 GROUPS ACTIVITY   <-- #

        #END OF ALL SWITCHES
        Write-Host "`nReport Data: `n"

        $ReportData
    }

    New-GrafanaSQLTable -dbServer $dbServer -ALL -DBName grafana -Credential $SQLCredential 

}
#endregion Table Creation

#region Update All Users data
Function Update-AllUsersData { 

    if (-not $AzureAppIDandSecret) {
        $AzureAppIDandSecret = Get-Credential -Message "Enter the ID and Secret for the Azure APP"
    }

    if (-not $SQLCredential) {
        $SQLCredential = Get-Credential -Message "Enter the SQL Credential"
    }

    Function Update-AllUsersData { 
        param(
            [String]$ReportRoot,
            [System.Management.Automation.PSCredential]$AzureAppIDandSecret,
            [String]$tenantID,
            [System.Management.Automation.PSCredential]$SQLCredential,
            [String]$DbServer,
            [String]$DBName
        )

        $ErrorTable = @()
        # Try on getting token

        Write-Log -Message "Start of $ReportRoot Function"
        try {
        
            $ReqTokenBody = @{
                Grant_Type    = "client_credentials"
                Scope         = "https://graph.microsoft.com/.default"
                client_Id     = $AzureAppIDandSecret.UserName
                Client_Secret = $AzureAppIDandSecret.GetNetworkCredential().Password
            } 

            $Tokparams = @{ 
                Uri             = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
                Method          = "POST"
                Body            = $ReqTokenBody
                UseBasicParsing = $True
            }
            Write-Log -Message "acquiring Graph API Token"
            $TokReqRes = Invoke-RestMethod @Tokparams

            $ReqHeader = @{
                Authorization = "Bearer $($TokReqRes.access_token)"
            }
            # Try on getting Graph Data
            try {
                Write-Log -Message "Gathering graph data for $reportRoot"
                $Graphparams = @{ 
                    Uri     = "https://graph.microsoft.com/v1.0/reports/$ReportRoot(period='D180')"
                    Method  = "Get"
                    Headers = $ReqHeader
                }

                try {
                    $GraphData = (Invoke-RestMethod @Graphparams) -replace "\xEF\xBB\xBF" | ConvertFrom-Csv
                }
                catch {
                    $Graphparams.Uri = "https://graph.microsoft.com/v1.0/reports/$ReportRoot"
                    $GraphData = (Invoke-RestMethod @Graphparams) -replace "\xEF\xBB\xBF" | ConvertFrom-Csv
                }
    
                if ($GraphData) {
                    Write-Log "Gathering Data for $reportRoot Success. Next: building query"
                
                    $columnHead = $GraphData | Get-Member -MemberType NoteProperty | Select-Object -expandProperty name

                    if (($columnHead -eq "User Principal Name")) {
                        $primKey = "User Principal Name"
                    }
                    elseif (($columnHead -eq "Owner Principal Name")) { 
                        $primKey = "Owner Principal Name"
                    } 
                    Else { 
                        $primKey = "Group ID"
                    }
                    $TableHeader = , $primKey + (($columnHead).foreach( { 
                                $_ #-replace '[^a-z0-9A-Z]', ''
                            }) -ne $primKey)

                    foreach ($Row in $GraphData) {

                        $UpdateClause = (($TableHeader[1..$TableHeader.Count].foreach( { 
                                        "[$_] = '$(($row.$_) -replace "'","''")',"
                                    }) -join "").trimEnd(',')) -replace ",", ",`n"

                        $SQLTableHeads = ((($TableHeader.ForEach( {
                                            "[$_]," 
                                        })) -join "").trimEnd(',')) -replace ",", ",`n"
    
                        $RowValues = (($TableHeader.foreach( { 
                                        "'$(($row.$_) -replace "'","''")',"
                                    }) -join "").trimEnd(',')) -replace ",", ",`n"
                    
                        # Lots of string manipulation here.
                        # It's just for debugging purposes and so the SQL query would look more readable. 
                        $query = @"
                        Update $ReportRoot Set
                            $UpdateClause
                        where [$($tableHeader[0])] = '$(
                            ($row.$($TableHeader[0])) -replace "'","''"
                        )'

                        IF @@ROWCOUNT=0
                            insert into $ReportRoot (
                                $SQLTableHeads       
                                ) 
                            values (
                                $RowValues
                                )
"@

                        # Actual inserting of the SQL query to DB
                        try {
                            write-log -message "Inserting $($row.$primKey)"
                            $SQLParams = @{ 
                                ServerInstance = $DbServer
                                Database       = $DBName
                                Credential     = $SQLCredential
                                Query          = $query
                                Verbose        = $true
                                AbortOnError   = $true
                            }
                            Invoke-Sqlcmd @SQLParams
                            #Uncomment this variable if you want to see the actual query string to DB. 
                            #$query
                        
                        }
                        catch {
                            $ErrorTable += [PSCustomObject]@{
                                Entity = $Row.$primKey
                                Table  = $ReportRoot
                            }
                            Write-Log -Message "Error on $($row.$primKey)" -level ERROR
                            Write-Log  -Level ERROR -Message "Last Error Message: $($Error[0].Exception.Message)"
                        }
                    } 
                }
            }
            # Catch on getting Graph Data
            catch {
                Write-Log  -Level ERROR -Message "Error on gathering data for $ReportRoot"
                Write-Log  -Level ERROR -Message "Last Error Message: $($Error[0].Exception.Message)"
                break
            }
        
        }
        #Catch on getting token
        catch { 
            Write-Log -Message "Error on acquiring graph Token" -Level ERROR
            Write-Log  -Level ERROR -Message "Last Error Message: $($Error[0].Exception.Message)"
            break
        }

        # Clean up on data that are no longer existing in Azure AD.
        Write-Log "Removing users in SQL that are no longer in Graph"
    
        $SQLParams = @{ 
            ServerInstance = $DbServer
            Database       = $DBName
            Credential     = $SQLCredential
            Verbose        = $true
            AbortOnError   = $true
        }
    
        try {
            $usersInSQL = Invoke-Sqlcmd @SQLParams -Query ("SELECT [$primKey] FROM $ReportRoot") |
                Select-Object -ExpandProperty $primKey
            $usersInGraph = $GraphData | Select-Object -ExpandProperty $primKey

            $DeletedinGraph = compare-object -ReferenceObject $usersInGraph -DifferenceObject $usersInSQL |
                Where-Object SideIndicator -EQ "=>" | Select-Object -ExpandProperty InputObject

            if ($DeletedinGraph) {
                $deleteWhere = (($DeletedinGraph.ForEach( { 
                                "'$_' OR [$primKey] = "
                            })) -join "").TrimEnd(" OR [$primKey] = ")
            
                $query = @"
                DELETE from $reportRoot where [$primKey] = $deleteWhere
"@
                Invoke-Sqlcmd @SQLParams -Query $query
                Write-Log -Message "Deleted users in graph were deleted from SQL: $($DeletedinGraph -join ",")"
            }

            else { 
                Write-Log -Message "No users need to be deleted from SQL"
            }
        }
        catch {
            Write-Log -Level WARNING -Message "Failure on removing deleted users in Graph from SQL DB"
            Write-Log  -Level WARNING -Message "Last Error Message: $($Error[0].Exception.Message)"
        }
    
        if ($errorTable) {
            Write-Log -Message ("Error Table: " + $($ErrorTable | OUT-STRING))
        }

        Wait-Logging
        Write-Log -Message "End of $ReportRoot Function"
    } # END OF FUNCTION

    $SQLParams = @{ 
        DBName              = $DBName
        AzureAppIDandSecret = $AzureAppIDandSecret
        tenantID            = $tenantID
        SQLCredential       = $SQLCredential
        DbServer            = $DbServer
    } 

    @(
        "getTeamsDeviceUsageUserDetail"
        "getTeamsUserActivityUserDetail"
        "getEmailActivityUserDetail"
        "getEmailAppUsageUserDetail"
        "getMailboxUsageDetail"
        "getOffice365ActivationsUserDetail"
        "getOffice365ActiveUserDetail"
        "getOffice365GroupsActivityDetail"
        "getOneDriveActivityUserDetail"
        "getOneDriveUsageAccountDetail"
        "getSharePointActivityUserDetail"
        "getSharePointSiteUsageDetail"
        "getSkypeForBusinessActivityUserDetail"
    
    )[0].foreach( {
            Update-AllUsersData @SQLparams -ReportRoot $_
        })

}
#endregion Update all users data

#region Update-SingleData
Function Update-SingleData { 
    
    Function Update-SingleRowData { 
        param(
            [String]$ReportRoot,
            [System.Management.Automation.PSCredential]$AzureAppIDandSecret,
            [String]$tenantID,
            [System.Management.Automation.PSCredential]$SQLCredential,
            [String]$DbServer,
            [String]$DBName,
            [Array]$TableHeader
        )

        $ErrorTable = @()
        # Try on getting token

        Write-Log -Message "Start of $reportRoot Function"
        try {
        
            $ReqTokenBody = @{
                Grant_Type    = "client_credentials"
                Scope         = "https://graph.microsoft.com/.default"
                client_Id     = $AzureAppIDandSecret.UserName
                Client_Secret = $AzureAppIDandSecret.GetNetworkCredential().Password
            } 

            $Tokparams = @{ 
                Uri             = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
                Method          = "POST"
                Body            = $ReqTokenBody
                UseBasicParsing = $True
            }
            $TokReqRes = Invoke-RestMethod @Tokparams

            $ReqHeader = @{
                Authorization = "Bearer $($TokReqRes.access_token)"
            }
            # Try on getting Graph Data
            try {
                Write-Log -Message "Gathering graph data for $reportRoot"
                $Graphparams = @{ 
                    Uri     = "https://graph.microsoft.com/v1.0/reports/$ReportRoot(period='D180')"
                    Method  = "Get"
                    Headers = $ReqHeader
                }
                try {
                    $GraphData = (Invoke-RestMethod @Graphparams) -replace "\xEF\xBB\xBF" | ConvertFrom-Csv
                }
                catch {
                    $Graphparams.Uri = "https://graph.microsoft.com/v1.0/reports/$ReportRoot"
                    $GraphData = (Invoke-RestMethod @Graphparams) -replace "\xEF\xBB\xBF" | ConvertFrom-Csv
                }
    
                if ($GraphData) {

                    #####################
                    # WRITE LOGICS HERE #
                    #####################

                    $TableHeader = ($GraphData | Get-Member -MemberType NoteProperty | Select-Object -expandProperty name).foreach( { 
                            $_ #-replace '[^a-z0-9A-Z\s]', ''
                        })

                    $RowValues = foreach ($row in $Graphdata) { 

                        $res = (($(
                                    $TableHeader.ForEach( { 
                                            "'$($row.$_)',"
                                        })
                                ) -join "").TrimEnd(","))
                        "($res);"
                    }

                    $RowValues = (($RowValues -join "").TrimEnd(";")) -replace ";", ",`n"

                    $sqlHead = ((($TableHeader).ForEach( { 
                                    "[$_],"
                                })) -join "").trimend(",")

                    $sqlHead = "($sqlHead)"

                    $query = @"
                    delete $ReportRoot;
                    Insert into $ReportRoot $sqlHead
                        VALUES
                        $rowValues
"@

                    try {
                        $SQLParams = @{ 
                            ServerInstance = $DbServer
                            Database       = $DBName
                            Credential     = $SQLCredential
                            Verbose        = $true
                            AbortOnError   = $true
                        }
                        Write-Log -message "Inserting current value to Database"
                        Invoke-Sqlcmd @SQLParams -query $query
                        Write-Log -Message "Success"
                        # Uncomment this variable if you want to see the actual query string to DB. 
                        # $query
                    }
                    catch {
                        $ErrorTable += [PSCustomObject]@{
                            Table  = $ReportRoot
                            Result = "Fail"
                        }
                        Write-Log -Message "Error on $($row.'User Principal Name')" -level ERROR
                        Write-Log  -Level ERROR -Message "Last Error Message: $($Error[0].Exception.Message)"
                    }
                }
            }
            # Catch on getting Graph Data
            catch {
                Write-Log  -Level ERROR -Message "Error on gathering data for $ReportRoot"
                Write-Log  -Level ERROR -Message "Last Error Message: $($Error[0].Exception.Message)"
                break
            }
        
        }
        #Catch on getting token
        catch { 
            Write-Log -Message "Error on acquiring graph Token" -Level ERROR
            Write-Log  -Level ERROR -Message "Last Error Message: $($Error[0].Exception.Message)"
            break
        }

    
    
        if ($errorTable) {
            Write-Log -Message ("Error Table: " + $($ErrorTable | OUT-STRING))
        }
        Wait-Logging
        Write-Log -Message "End of $reportRoot Function"
    }# END OF FUNCTION

    $sqlParams = @{ 
        DBName              = $DBName
        AzureAppIDandSecret = $AzureAppIDandSecret
        tenantID            = $tenantID
        SQLCredential       = $SQLCredential
        DbServer            = $DbServer
    } 

    @(
        "getTeamsDeviceUsageDistributionUserCounts"
        "getEmailAppUsageAppsUserCounts"
        "getEmailAppUsageVersionsUserCounts"
        "getOffice365ActivationCounts"
        "getOffice365ActivationsUserCounts"
        "getOffice365ServicesUserCounts"
    
    ).foreach( {
            Update-SingleRowData @sqlParams -ReportRoot $_
        })

}
#endregion Update-SIngleData

#region Update-UsageCountData
Function Update-UsageCountData { 
    
    Function Update-UsageCountData { 
        param(
            [String]$ReportRoot,
            [System.Management.Automation.PSCredential]$AzureAppIDandSecret,
            [String]$tenantID,
            [System.Management.Automation.PSCredential]$SQLCredential,
            [String]$DbServer,
            [String]$DBName
        )

        $ErrorTable = @()
        # Try on getting token
        Write-Log -Message "Start of $ReportRoot Function"
        try {
        
            $ReqTokenBody = @{
                Grant_Type    = "client_credentials"
                Scope         = "https://graph.microsoft.com/.default"
                client_Id     = $AzureAppIDandSecret.UserName
                Client_Secret = $AzureAppIDandSecret.GetNetworkCredential().Password
            } 

            $Tokparams = @{ 
                Uri             = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
                Method          = "POST"
                Body            = $ReqTokenBody
                UseBasicParsing = $True
            }
            Write-Log -Message "acquiring Graph API Token"
            $TokReqRes = Invoke-RestMethod @Tokparams

            $ReqHeader = @{
                Authorization = "Bearer $($TokReqRes.access_token)"
            }
            # Try on getting Graph Data
            try {

                $Graphparams = @{ 
                    Uri     = "https://graph.microsoft.com/v1.0/reports/$ReportRoot(period='D180')"
                    Method  = "Get"
                    Headers = $ReqHeader
                }
                Write-Log -Message "Gathering graph data for $reportRoot"

                $GraphData = (Invoke-RestMethod @Graphparams) -replace "\xEF\xBB\xBF" | ConvertFrom-Csv

                # Exceptions on the formatting of the result. These don't follow the normal report format. These catches are to reformat them.
                if ($ReportRoot -eq "getOneDriveUsageStorage") {
                    $GraphData = $GraphData | Where-Object { $_.'Site Type' -EQ "OneDrive" }
                }
 
                $TableHeader = , "Report Date" + (($GraphData | Get-Member -MemberType NoteProperty | Select-Object -expandProperty name).foreach( { 
                            $_ #-replace '[^a-z0-9A-Z\s]', ''
                        }) -ne "Report Date")

                function InsertTo-DB { 
                
                    # This function is being called by the logics below. It inserts the data to DB
                    # Start of creating SQL Query strings

                    $RowValues = foreach ($row in $ToInserObj) { 

                        $res = (($(
                                    $TableHeader.ForEach( { 
                                            "'$($row.$_)',"
                                        })
                                ) -join "").TrimEnd(","))
                        "($res);"
                    }
                    $RowValues = (($RowValues -join "").TrimEnd(";")) -replace ";", ",`n"

                    $sqlHead = ((($TableHeader).ForEach( { 
                                    "[$_],"
                                })) -join "").trimend(",")

                    $sqlHead = "($sqlHead)"

                    $Query = @"
                    Insert into $ReportRoot $sqlHead 
                        VALUES
                        $RowValues
"@
                    try {
                        Write-Log -Message "Inserting new data to SQL Database"
                        Invoke-Sqlcmd @SQLParams -Query $Query
                        Write-Log -Message "Inserting new data to DB SUCCESSFUL"
                        write-log -message "$($ToInserObj.count) rows inserted"

                    }
                    catch {

                        $ErrorTable += [PSCustomObject]@{
                            Entity = "DateRange"
                            Table  = $ReportRoot
                            Result = "Fail"
                        }
                        write-log -message "Failed in inserting to table $ReportRoot"

                    }
                } #End of Function Insert to DB

                try {
                    $SQLParams = @{ 
                        ServerInstance = $DbServer
                        Database       = $DBName
                        Credential     = $SQLCredential
                        Verbose        = $true
                        AbortOnError   = $true
                    }

                    # Gathers the existing data in the database. I only care about 180 days since this will be compared
                    # To the data in Graph which can also only go up to last 180 days. 
                    Write-Log -Message "Gathering exiting last 180 rows of data in database"
                    $Top180Date = Invoke-Sqlcmd @SQLParams -Query "Select Top (180) [Report Date] from $ReportRoot order by [Report Date] Desc"
                
                    $ToInserObj = @()
                    if ($Null -ne $Top180Date) {

                        Write-Log -Message "Processing/parsing data to insert"
                        $RefObjGraphdata = $GraphData | Select-Object -expand "Report Date"
                        $DiffObjSQLData = $Top180Date |
                            Select-Object @{ Name = "Report Date"; E = { Get-Date $_.'Report Date' -Format "yyyy-MM-dd" } } |
                            Select-Object -expand "Report Date"
                    
                        # Doing a comparison of what's not in the database. Whatever that will be caught, it will be
                        # Forwarded to insert function to by saved in DB.
                        $MissingDates = compare-object -ReferenceObject $RefObjGraphdata -DifferenceObject $DiffObjSQLData |
                            Where-Object { $_.SideIndicator -eq "<=" }

                        if ($null -ne $MissingDates) { 

                            foreach ($item in $MissingDates) {
                                $ToInserObj += $GraphData | Where-Object { $_.'Report Date' -eq $item.InputObject }

                            }
                            #cALL INSERrt to db function
                            InsertTo-DB
                        }
                        else { 
                            Write-Log -Message "No new Data to Add"
                        }

                    }

                    #SQL db is empty. Will insert all Graph Data
                    else { 
                        $ToInserObj = $GraphData
                        InsertTo-DB
                    }

                    #############
                }

                # Catch on getting current DB Data
                catch {
                    write-log -message "Unable to query existing rows in table $ReportRoot"
                    Write-Log  -Level ERROR -Message "Last Error Message: $($Error[0].Exception.Message)"
                    break
                }
            }
            # Catch on getting Graph Data
            catch {
                Write-Log  -Level ERROR -Message "Error on gathering data for $ReportRoot"
                Write-Log  -Level ERROR -Message "Last Error Message: $($Error[0].Exception.Message)"
                break
            }        
        }
        #Catch on getting token
        catch { 
            Write-Log -Message "Error on acquiring graph Token" -Level ERROR
            Write-Log  -Level ERROR -Message "Last Error Message: $($Error[0].Exception.Message)"
            break
        }

    
        Write-Log -Message "End of $ReportRoot Function"
        if ($errorTable) {
            Write-Log -Message ("Error Table: " + $($ErrorTable | OUT-STRING))
        }
        Wait-Logging
    }# END OF FUNCTION

    $sqlParams = @{ 
        DBName              = $DBName
        AzureAppIDandSecret = $AzureAppIDandSecret
        tenantID            = $tenantID
        SQLCredential       = $SQLCredential
        DbServer            = $DbServer
    } 

    @(
        "getTeamsDeviceUsageUserCounts"
        "getTeamsUserActivityCounts"
        "getTeamsUserActivityUserCounts"
        "getEmailActivityCounts"
        "getEmailActivityUserCounts"
        "getEmailAppUsageUserCounts"
        "getMailboxUsageMailboxCounts"
        "getMailboxUsageQuotaStatusMailboxCounts"
        "getMailboxUsageStorage"
        "getOffice365ActiveUserCounts"
        "getOffice365GroupsActivityCounts"
        "getOffice365GroupsActivityGroupCounts"
        "getOffice365GroupsActivityStorage"
        "getOffice365GroupsActivityFileCounts"
        "getOneDriveActivityUserCounts"
        "getOneDriveActivityFileCounts"
        "getOneDriveUsageAccountCounts"
        "getOneDriveUsageFileCounts"
        "getOneDriveUsageStorage"
        "getSharePointActivityFileCounts"
        "getSharePointActivityUserCounts"
        "getSharePointActivityPages"
        "getSharePointSiteUsageFileCounts"
        "getSharePointSiteUsageSiteCounts"
        "getSharePointSiteUsageStorage"
        "getSharePointSiteUsagePages"
    
    ).foreach( {
            Update-UsageCountData @SQLparams -ReportRoot $_
        
        })

}
#endregion Update-UsageCountData

# One time use only. Only during table creating. 
#Create-Table 

Update-AllUsersData
Update-SingleData
Update-UsageCountData
