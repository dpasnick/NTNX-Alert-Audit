###################################################################################################
#
# Created: 08/02/19
# Modified: 02/02/22
#
# Description: Uses the standard Nutanix Prism API calls (v2) along with LCM API calls in order to
#              pull alert information from individual clusters and export the informaiton in an 
#              Excel file.
#
# Requires: - Common account (local or directory) that will allow login to multiple clusters with 
#             the same credentials.
#           - CSV input file nust one columne with heading Ext_IP (the external/VIP address of the cluster)
#
# Note: An Excel spreadsheet will be generated and placed in the same directory as script.
#
###################################################################################################

$ErrorActionPreference = "silentlycontinue"


$fileCSV = $(get-location).Path + "\Nutanix-Cluster_Lookup.csv"


############################################
## Do not change anything below this line ##
############################################

# Initial Screen Messaging
clear;
Write-Host $dateToday
Write-Host "This script will use '$fileCSV' to pull cluster information.`n" -ForegroundColor Yellow

# Username input - either local account or AD account if cluster is configured
$username = Read-Host -Prompt 'Enter cluster(s) administrative username.'

# Password input - either local account or AD account if cluster is configured
$securePassword = Read-Host -Prompt 'Enter cluster(s) administrative password.' -AsSecureString
$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))

$myTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Central Standard Time")


### AUTHORIZATION ###
$Header = @{"Authorization" = "Basic "+[System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($username+":"+$password ))}
Add-Type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
### End Authorization ###


### INVOKE REST Function ###
function remoteConnect ($uri) {    
    Write-Output $uri
    Invoke-RestMethod -Uri $uri -Header $Header
    return
}
### END INVOKE REST Function ###


### Convert Epoc Date Function ###
function convertDate ($usec) {
    ([System.TimeZoneInfo]::ConvertTimeFromUtc((New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0).AddSeconds([math]::Floor($usec/1000000)),$myTimeZone)).ToString()
    return
}
### End Convert Epoc Date Function ###


### Alert Messages - Add Context Info ###
function alertContext ($alertUpdate, $alertEntity) {
    $message = $alertUpdate.$alertEntity
    for ($i = 0; $i -lt ($alertUpdate.context_types).Count; $i++) {
        if (-not [System.String]::IsNullOrEmpty($alertUpdate.context_types[$i])) {
            $message = $message.Replace("{$($alertUpdate.context_Types[$i])}", $alertUpdate.context_values[$i]);
        }
    }
    return $message
}
### End Alert Messages - Add Context Info ###


### Loop through clutsers and Alerts ###
$csv = Import-Csv "$fileCSV"

# Attempt to connect to all PC instances and pull stats
try {

# Run through each Prism Element Cluster Instance and pull API payloads.
foreach ($cluster in $csv) {
    $ntnxCluster = $($cluster.Ext_IP)

    Write-Host "`nConnecting to Nutanix Cluster Instance $ip ...`n" -foregroundColor Yellow
    Write-Host "  - Triggering REST API for $ntnxCluster..." -foregroundColor Green

    $uriAlerts = "https://$($ntnxCluster):9440/PrismGateway/services/rest/v2.0/alerts/?resolved=false&acknowledged=false"
    $uriCluster = "https://$($ntnxCluster):9440/PrismGateway/services/rest/v2.0/cluster/"
    $clusterName = (remoteConnect $uriCluster).name
    $alertList = remoteConnect $uriAlerts

    ForEach ($iAlert in $alertList.entities) {
        $alertItem = [PSCustomObject]@{
            Cluster_Name = $clusterName.ToUpper()
            Alert_ID = $iAlert.alert_type_uuid
            Alert_Severity = $iAlert.severity
            Creation_Time = (convertDate $iAlert.created_time_stamp_in_usecs)
            Last_Occurance = (convertDate $iAlert.last_occurrence_time_stamp_in_usecs)
            Alert_Title = (alertContext $iAlert "alert_title")
            Alert_Message =  (alertContext $iAlert "message")
        }
        $alertItems+=$alertItem
    }
}

} catch {

    Write-Host "`n`n*************************************************************************" -ForegroundColor Red
    Write-Host "Error Type: " $_.Exception.Message -ForegroundColor Red
    Write-Host "Error Line: " $_.InvocationInfo.ScriptLineNumber -ForegroundColor Red
    Write-Host "Failed For: " $ntnxCluster -ForegroundColor Red
    Write-Host "*************************************************************************`n`n" -ForegroundColor Red

}


### Launch Excel Application ###
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $True
$workBook = $excelApp.Workbooks.Add()
### End Launch Excel Application ###

### Export Cluster Alerts to Excel ###
$workSheet = $workBook.Worksheets.Item(1)
$workSheet.Rows.HorizontalAlignment = -4131 
$workSheet.Rows.Font.Size = 10
$workSheet.Name =  "Alerts"
$row = $col = 1
$alertXLHead = ("Cluster_Name","Alert_ID","Alert_Severity","Creation_Time","Last_Occurance","Alert_Title","Alert_Message")
$alertXLHead | %( $_  ){ $workSheet.Cells.Item($row,$col) = $_ ; $col++ }
$workSheet.Rows.Item(1).Font.Bold = $True
$workSheet.Rows.Item(1).HorizontalAlignment = -4108

$i = 0; $row++; $col = 1
FOREACH( $alertItemE in $alertItems ){ 
    $i = 0
    DO{ 
        $workSheet.Cells.Item($row,$col) = $alertItemE.($alertXLHead[$i])
        $col++
        $i++ 
    }UNTIL($i -ge $alertXLHead.Count)
    $row++; $col=1
} 
$workSheet.UsedRange.EntireColumn.AutoFit()
### End Export Cluster Alerts to Excel ###

### Save Excel Workbook ###
$Date = Get-Date
$Today = (Get-Date).toshortdatestring().Replace("/","-")
$filepath = $(get-location).Path + "\Nutanix_Alert_Report-$Today.xlsx"
$excelApp.DisplayAlerts = $False
$workBook.SaveAs($filepath)
$excelApp.Quit()
$filepath
### End Save Excel Workbook ###

# Cleanup
Remove-Variable username -ErrorAction SilentlyContinue
Remove-Variable securePassword -ErrorAction SilentlyContinue
Remove-Variable password -ErrorAction SilentlyContinue
# End Cleanup
