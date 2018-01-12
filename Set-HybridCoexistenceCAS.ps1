[CmdletBinding(DefaultParameterSetName='AllServers')]
param (
    [Parameter(
        Mandatory,
        ParameterSetName='AllServers'
    )]
    [Parameter(
        Mandatory,
        ParameterSetName='ServersOnly'
    )]
    [Parameter(
        Mandatory,
        ParameterSetName='ServersByVersion'
    )]
    [String]$Namespace,

    [Parameter(
        Mandatory,
        ParameterSetName='AllServers'
    )]
    [Parameter(
        Mandatory,
        ParameterSetName='ServersOnly'
    )]
    [Parameter(
        Mandatory,
        ParameterSetName='ServersByVersion'
    )]
    [String]$Hostname,

    [Parameter(
        Mandatory,
        ParameterSetName='ServersByVersion'
    )]
    [ValidateSet("2010","2013","2016")]
    [String]$Version,

    [Parameter(
        Mandatory,
        ValueFromPipeline,
        ParameterSetName='ServersOnly'
    )]
    [String[]]$Servers,

    [Alias("AutoD")]
    [Parameter(
        Mandatory=$False,
        ParameterSetName='AllServers'
    )]
    [Parameter(
        Mandatory=$False,
        ParameterSetName='ServersOnly'
    )]
    [Parameter(
        Mandatory=$False,
        ParameterSetName='ServersByVersion'
    )]
    [Switch]$AutoDiscover
)

Begin {
    $FQDN = "${Hostname}.${Namespace}"
    $AUTO = "https://autodiscover.${Namespace}/Autodiscover/Autodiscover.xml"
    $OWA = "https://${FQDN}/owa"
    $ECP = "https://${FQDN}/ecp"
    $OAB = "https://${FQDN}/oab"
    $EWS = "https://${FQDN}/EWS/Exchange.asmx"
    $EAS = "https://${FQDN}/Microsoft-Server-ActiveSync"
    $MAPI = "https://${FQDN}/mapi"
    $CAS = $FQDN

    $ServerList = New-Object System.Collections.Generic.List[System.Object]
}

Process{
    If($Servers) {
        ForEach($Server in $Servers) {
            $ServerList.Add($Server) | Out-Null
        }
        $ServerList = $ServerList | Get-ExchangeServer
    } ElseIf ($Version) {
        Switch ($Version) {
            "2010" { $AdmDispVer = "*14.*" }
            "2013" { $AdmDispVer = "*15.0*" }
            "2016" { $AdmDispVer = "*15.1*"}
        }
        $ServerList = Get-ExchangeServer | Where-Object {$_.AdminDisplayVerion -Like $AdmDispVer}
    } Else {
        $ServerList = Get-ExchangeServer
    }
}

End {
    $ServerList | ForEach-Object {
        $Bypass = "https://$($_.Fqdn)/EWS/Exchange.asmx"
        $_ | Get-OwaVirtualDirectory | Set-OwaVirtualDirectory -ExternalUrl $OWA -InternalUrl $OWA
        $_ | Get-EcpVirtualDirectory | Set-EcpVirtualDirectory -ExternalUrl $ECP -InternalUrl $ECP
        $_ | Get-OabVirtualDirectory | Set-OabVirtualDirectory -ExternalUrl $OAB -InternalUrl $OAB -RequireSSL $True
        $_ | Get-ActiveSyncVirtualDirectory | Set-ActiveSyncVirtualDirectory -ExternalUrl $EAS -InternalUrl $EAS
        $_ | Get-WebServicesVirtualDirectory | Set-WebServicesVirtualDirectory -ExternalUrl $EWS -InternalUrl $EWS -InternalNLBBypassUrl  $Bypass -Force
        if($AutoDiscover) {$_ | Get-ClientAccessServer | Set-ClientAccessServer -AutoDiscoverServiceInternalUri $AUTO}
        $_ | Get-OutlookAnywhere | Set-OutlookAnywhere -InternalHostname $CAS -ExternalHostname $CAS -InternalClientsRequireSsl:$True -InternalClientAuthenticationMethod Ntlm -ExternalClientsRequireSsl:$True -ExternalClientAuthentication Negotiate -IISAuthentication Basic,Ntlm,Negotiate
        $_ | Get-MapiVirtualDirectory | Set-MapiVirtualDirectory -ExternalUrl $MAPI -InternalUrl $MAPI
    }

    $ServerList | ForEach-Object {
        Write-Output "Server: $_"
        Write-Output "OWA"
        $_ | Get-OwaVirtualDirectory -ADPropertiesOnly | Format-List *lUrl
        Write-Output "ECP"
        $_ | Get-EcpVirtualDirectory -ADPropertiesOnly | Format-List *Url
        Write-Output "OAB"
        $_ | Get-OabVirtualDirectory -ADPropertiesOnly | Format-List *Url
        Write-Output "Exchange ActiveSync"
        $_ | Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | Format-List *lUrl
        Write-Output "EWS"
        $_ | Get-WebServicesVirtualDirectory -ADPropertiesOnly | Format-List *Url
        Write-Output "AutoDiscover"
        $_ | Get-ClientAccessServer | Format-List AutoDiscoverServiceInternalUri
        Write-Output "Outlook Anywhere"
        $_ | Get-OutlookAnywhere -ADPropertiesOnly | Format-List *lHostname*,*eSsl*,*Method*
        Write-Output "MAPI"
        $_ | Get-MapiVirtualDirectory -ADPropertiesOnly | Format-List *lUrl
    }
}