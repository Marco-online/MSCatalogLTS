function Invoke-DownloadFile {
    [CmdLetBinding()]
    param (
        [uri] $Uri,
        [string] $Path
    )
    
    try {
        if (Test-Path $Path) {
            $Hash = Get-FileHash -Path $Path -Algorithm SHA1
            if ($Path -match "$($Hash.Hash)\.msu$") {
                Write-Verbose "File already exists"
                return
            }
        }
        
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