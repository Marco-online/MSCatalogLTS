function Invoke-DownloadFile {
    [CmdLetBinding()]
    param (
        [uri] $Uri,
        [string] $Path
    )
    
    try {
        Set-TempSecurityProtocol

        $WebClient = [System.Net.WebClient]::new()
        $WebClient.DownloadFile($Uri, $Path)
        $WebClient.Dispose()
    } catch {
        $Err = $_
        if ($WebClient) {
            $WebClient.Dispose()
        }
        throw $Err
    }
    Set-TempSecurityProtocol -ResetToDefault
}