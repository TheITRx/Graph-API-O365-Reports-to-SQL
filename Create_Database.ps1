<#
.SYNOPSIS
    . Creates the SQL server, database, user account, and firewall exception where you need to managed the DB
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
    General notes
#>

$SubID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxx"
$Location = "northcentralus"
$ResourceGroup = "xxxrg"
$SQLAdminCred = Get-Credential -Message "Enter the SQL Admin credentials you want to set"
$SQLSRVName = "xx"
$SQLDBName = "xx"
$SQLDBTier = "Basic"
$sqlClientPIP = "x.x.x.x"

# End of static variables

$SQLAdminUN = $SQLAdminCred.UserName
$SQLAdminPW = $SQLAdminCred.GetNetworkCredential().Password

az login 

write-host "Setting the default subscription"
az account set --subscription $SubID

write-host "Creating the resource Group"
az group create `
    --location $Location `
    --name $ResourceGroup `
    --location $Location `
    --verbose

write-host "Creating the SQL Server"
az sql server create `
    --admin-user $SQLAdminUN `
    --admin-password $SQLAdminPW `
    --name $SQLSRVName `
    --resource-group $ResourceGroup `
    --location $Location `
    --verbose

write-host "Creating the Database"
az sql db create `
    --name $SQLDBName `
    --resource-group $ResourceGroup `
    --server $SQLSRVName `
    --tier $SQLDBTier `
    --verbose


write-host "Creating the firewall Rule"
az sql server firewall-rule create `
    --name grafanaapp `
    --server $SQLSRVName `
    --resource-group $ResourceGroup `
    --start-ip-address $sqlClientPIP `
    --end-ip-address $sqlClientPIP `
    --verbose
        