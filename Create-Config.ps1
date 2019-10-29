$config = New-Object -TypeName PSObject -Property @{ }

$azureIdandKey = Get-Credential -Message "Azure App ID and Key"
$sqlCredential = Get-Credential -Message "Enter SQL Credential"

$tenantID = Read-Host -Prompt "Enter the Tenant name or ID:"
$DBName = Read-Host -Prompt "Enter the DB Name:"
$DBServer = Read-Host -Prompt "Enter the DB Server:"

$config | Add-Member -MemberType NoteProperty -Name AzureAppID -Value ($azureIdandKey.UserName)
$config | Add-Member -MemberType NoteProperty -Name AzureAppKey -Value ($azureIdandKey.password | ConvertFrom-SecureString)

$config | Add-Member -MemberType NoteProperty -Name sqlCredentialUserName -Value ($sqlCredential.UserName)
$config | Add-Member -MemberType NoteProperty -Name sqlCredentialPassword -Value ($sqlCredential.password | ConvertFrom-SecureString)

$config | Add-Member -MemberType NoteProperty -Name TenantID -Value $tenantID
$config | Add-Member -MemberType NoteProperty -Name DBName -Value $DBName
$config | Add-Member -MemberType NoteProperty -Name DBServer -Value $DBServer

$config | ConvertTo-Json -Depth 3 | Out-File $PSScriptRoot\Configuration.json