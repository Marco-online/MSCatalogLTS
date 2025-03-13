function Get-UpdateLinks {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0
        )]
        [String] $Guid
    )

    $Post = @{size = 0; UpdateID = $Guid; UpdateIDInfo = $Guid} | ConvertTo-Json -Compress
    $Body = @{UpdateIDs = "[$Post]"}
    
    $Params = @{
        Uri = "https://www.catalog.update.microsoft.com/DownloadDialog.aspx"
        Body = $Body
        ContentType = "application/x-www-form-urlencoded"
        UseBasicParsing = $true
    }

    $DownloadDialog = Invoke-WebRequest @Params
    $Links = $DownloadDialog.Content -replace "www.download.windowsupdate", "download.windowsupdate"

    $Regex = "downloadInformation\[0\]\.files\[\d+\]\.url\s*=\s*'([^']*kb(\d+)[^']*)'"
    $DownloadMatches = [regex]::Matches($Links, $Regex)

    if ($DownloadMatches.Count -eq 0) {
        $RegexFallback = "downloadInformation\[0\]\.files\[0\]\.url\s*=\s*'([^']*)'"
        $DownloadMatches = [regex]::Matches($Links, $RegexFallback)
    }

    if ($DownloadMatches.Count -eq 0) {
        return $null
    }
    
    $KbLinks = foreach ($Match in $DownloadMatches) {
        [PSCustomObject]@{
            URL = $Match.Groups[1].Value
            KB  = if ($Match.Groups.Count -gt 2 -and $Match.Groups[2].Success) { [int]$Match.Groups[2].Value } else { 0 }
        }
    }
    
    return $KbLinks | Sort-Object KB -Descending
}