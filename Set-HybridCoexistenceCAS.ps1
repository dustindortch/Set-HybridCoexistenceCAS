[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String]$Namespace,
    [Parameter(Mandatory=$False)]
    [ValidateSet("2013","2016")]
    [String]$Version,
    [Parameter(Mandatory=$False)]
    [Switch]$AutoD
)

Switch ($Version) {
    "2013" { $AdmDispVer = "*15.0*" }
    "2016" { $AdmDispVer = "*15.1*"}
    Default { $AdmDispVer = "*15.*"}
}

$AutoD = "https://autodiscover.${Namespace}/Autodiscover/Autodiscover.xml"
$OWA = "https://${Namespace}/owa"
$ECP = "https://${Namespace}/ecp"
$OAB = "https://${Namespace}/oab"
$EWS = "https://${Namespace}/ews"
$EAS = "https://${Namespace}/Microsoft-Server-ActiveSync"
$MAPI = "https://${Namespace}/mapi"
$CAS = $Namespace

$Servers = Get-ExchangeServer | Where-Object {$_.AdminDisplayVerion -Like $AdmDispVer}

$Servers | ForEach-Object {
    $Bypass = ($_ | Get-WebServicesVirtualDirectory).InternalUrl
    $_ | Get-OwaVirtualDirectory | Set-OwaVirtualDirectory -ExternalUrl $OWA -InternalUrl $OWA
    $_ | Get-EcpVirtualDirectory | Set-EcpVirtualDirectory -ExternalUrl $ECP -InternalUrl $ECP
    $_ | Get-OabVirtualDirectory | Set-OabVirtualDirectory -ExternalUrl $OAB -InternalUrl $OAB -RequireSSL $True
    $_ | Get-ActiveSyncVirtualDirectory | Set-ActiveSyncVirtualDirectory -ExternalUrl $EAS -InternalUrl $EAS
    $_ | Get-WebServicesVirtualDirectory | Set-WebServicesVirtualDirectory -ExternalUrl $EWS -InternalUrl $EWS -InternalBypassUrl  $Bypass.OriginalString
    if($AutoD) {$_ | Get-ClientAccessServer | Set-ClientAccessServer -AutoDiscoverServiceInternalUri $AutoD}
    $_ | Get-OutlookAnywhere | Set-OutlookAnywhere -InternalHostname $CAS -ExternalHostname $CAS
    $_ | Get-MapiVirtualDirectory | Set-MapiVirtualDirectory -ExternalUrl $MAPI -InternalUrl $MAPI
}

$Servers | ForEach-Object {
    Write-Output "Server: $_"
    Write-Output "OWA"
    $_ | Get-OwaVirtualDirectory | FL *lUrl
    Write-Output "ECP"
    $_ | Get-EcpVirtualDirectory | FL *Url
    Write-Output "OAB"
    $_ | Get-OabVirtualDirectory | FL *Url
    Write-Output "Exchange ActiveSync"
    $_ | Get-ActiveSyncVirtualDirectory | FL *lUrl
    Write-Output "EWS"
    $_ | Get-WebServicesVirtualDirectory | FL *Url
    Write-Output "AutoDiscover"
    $_ | Get-ClientAccessServer | FL AutoDiscoverServiceInternalUri
    Write-Output "Outlook Anywhere"
    $_ | Get-OutlookAnywhere | FL *lHostname*,*eSsl*,*Method*
    Write-Output "MAPI"
    $_ | Get-MapiVirtualDirectory | FL *lUrl
}

