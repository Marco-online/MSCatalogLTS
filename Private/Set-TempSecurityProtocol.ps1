function Set-TempSecurityProtocol {
    [CmdletBinding()]
    param (
        [switch] $ResetToDefault
    )

    if (($null -ne $Script:MSCatalogSecProt) -and $ResetToDefault) {
        [Net.ServicePointManager]::SecurityProtocol = $Script:MSCatalogSecProt
    } else {
        if ($null -eq $Script:MSCatalogSecProt) {
            $Script:MSCatalogSecProt = [Net.ServicePointManager]::SecurityProtocol
        }
        $Tls11 = [System.Net.SecurityProtocolType]::Tls11
        $Tls12 = [System.Net.SecurityProtocolType]::Tls12
        $CurrentProtocol = [Net.ServicePointManager]::SecurityProtocol
        $NewProtocol = $CurrentProtocol -bor $Tls11 -bor $Tls12
        [Net.ServicePointManager]::SecurityProtocol = $NewProtocol
    }
}