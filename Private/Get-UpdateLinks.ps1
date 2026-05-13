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

    # Build a lookup of SHA1 digests (base64) keyed by "downloadInfoIndex_fileIndex"
    $DigestRegex = "downloadInformation\[(\d+)\]\.files\[(\d+)\]\.digest\s*=\s*'([^']*)'"
    $DigestMatches = [regex]::Matches($Links, $DigestRegex)
    $DigestMap = @{}
    foreach ($DigestMatch in $DigestMatches) {
        $DigestKey = "$($DigestMatch.Groups[1].Value)_$($DigestMatch.Groups[2].Value)"
        $DigestMap[$DigestKey] = $DigestMatch.Groups[3].Value
    }

    $KbLinks = foreach ($Match in $DownloadMatches) {
        $InfoIdx = $Match.Groups[1].Value
        $FileIdx = $Match.Groups[2].Value
        $Url = $Match.Groups[3].Value

        # Try to extract KB number from the URL (if present)
        $KbNumber = 0
        if ($Url -match 'kb(\d+)') {
            $KbNumber = [int]$Matches[1]
        }

        $Sha1Base64 = ""
        $DigestKey = "${InfoIdx}_${FileIdx}"
        if ($DigestMap.ContainsKey($DigestKey)) {
            $Sha1Base64 = $DigestMap[$DigestKey]
        }

        [PSCustomObject]@{
            URL = $Url
            KB  = $KbNumber
            DownloadInfoIndex = [int]$InfoIdx
            FileIndex = [int]$FileIdx
            SHA1 = $Sha1Base64
        }
    }
    
    # Remove duplicates based on URL
    $UniqueLinks = $KbLinks | Group-Object -Property URL | ForEach-Object { $_.Group[0] }
    
    return $UniqueLinks | Sort-Object KB -Descending
}