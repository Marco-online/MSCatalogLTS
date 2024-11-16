function Invoke-DownloadFile {
    [CmdLetBinding()]
    param (
        [uri] $Uri,
        [string] $Path
    )
    
    try {      
        Set-TempSecurityProtocol
        
        $WebClient = [System.Net.WebClient]::new()
        Write-Verbose "Downloading file from $Uri to $Path"
        $WebClient.DownloadFile($Uri, $Path)
    } catch {
        throw
    } finally {
        if ($WebClient) {
            $WebClient.Dispose()
        }
        Set-TempSecurityProtocol -ResetToDefault
    }
}