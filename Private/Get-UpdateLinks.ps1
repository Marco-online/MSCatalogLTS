function Get-UpdateLinks {
    [CmdLetBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0
        )]
        [String] $Guid
    )

    $Post = @{size = 0; updateID = $Guid; uidInfo = $Guid} | ConvertTo-Json -Compress
    $Body = @{updateIDs = "[$Post]"}

    $Params = @{
        Uri = "https://www.catalog.update.microsoft.com/DownloadDialog.aspx"
        Method = "Post"
        Body = $Body
        ContentType = "application/x-www-form-urlencoded"
        UseBasicParsing = $true
    }
    $DownloadDialog = Invoke-WebRequest @Params
    $Links = $DownloadDialog.Content -replace "www.download.windowsupdate", "download.windowsupdate"
	$Regex = "downloadInformation\[0\]\.files\[0\]\.url\s*=\s*'([^']*)'"
   	$Links = [regex]::Matches($Links, $Regex)
   
    foreach ($Link in $Links) {
    Write-Output $Link.Groups[1].Value
        }
    } 