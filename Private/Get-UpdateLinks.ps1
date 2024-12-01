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
        Write-Warning "No download links found for the specified update."
        return $null
    }
    
    $KbLinks = foreach ($Match in $DownloadMatches) {
        [PSCustomObject]@{
            URL = $Match.Groups[1].Value
            KB  = [int]$Match.Groups[2].Value
        }
    }
    
    $Links = $KbLinks | Sort-Object KB -Descending
    $Links[0].URL
	}
    