
# Change these variables to your liking

$SubID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxx"
$Location = "northcentralus"
$ResourceGroup = "xxxrg"
$SQLAdminCred = Get-Credential -Message "Enter the SQL Admin credentials you want to set"
$SQLSRVName = "xx"
$SQLDBName = "xx"
$SQLDBTier = "Basic"
$GrafanaAppPIP = "x.x.x.x"

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
    --start-ip-address $GrafanaAppPIP `
    --end-ip-address $GrafanaAppPIP `
    --verbose
        