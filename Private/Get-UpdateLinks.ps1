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

    # NEW: Capture ALL downloadInformation arrays (not just [0])
    # This regex matches downloadInformation[ANY_INDEX].files[ANY_INDEX].url
    $Regex = "downloadInformation\[(\d+)\]\.files\[(\d+)\]\.url\s*=\s*'([^']*)'"
    $DownloadMatches = [regex]::Matches($Links, $Regex)

    if ($DownloadMatches.Count -eq 0) {
        Write-Verbose "No download links found in response."
        return $null
    }
    
    Write-Verbose "Found $($DownloadMatches.Count) download link(s) in response."
    
    $KbLinks = foreach ($Match in $DownloadMatches) {
        $Url = $Match.Groups[3].Value  # URL is in Group 3 now
        
        # Try to extract KB number from the URL (if present)
        $KbNumber = 0
        if ($Url -match 'kb(\d+)') {
            $KbNumber = [int]$Matches[1]
        }
        
        [PSCustomObject]@{
            URL = $Url
            KB  = $KbNumber
            DownloadInfoIndex = [int]$Match.Groups[1].Value
            FileIndex = [int]$Match.Groups[2].Value
        }
    }
    
    # Remove duplicates based on URL
    $UniqueLinks = $KbLinks | Group-Object -Property URL | ForEach-Object { $_.Group[0] }
    
    return $UniqueLinks | Sort-Object KB -Descending
}