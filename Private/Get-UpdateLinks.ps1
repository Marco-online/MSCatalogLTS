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
	$Regex = "downloadInformation\[0\]\.files\[\d+\]\.url\s*=\s*'([^']*kb(\d+)[^']*)'"
	$Matches = [regex]::Matches($Links, $Regex)

	$KbLinks = foreach ($Match in $Matches) {
		[PSCustomObject]@{
			URL = $Match.Groups[1].Value
			KB  = [int]$Match.Groups[2].Value
		}
	}

	$Links = $KbLinks | Sort-Object KB -Descending
	$Links[0].URL
	}
    