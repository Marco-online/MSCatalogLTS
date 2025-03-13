function Save-MSCatalogUpdate {
    param (
        [Parameter(
            Position = 0, 
            ParameterSetName = "ByObject")]
        [Object] $Update,

        [Parameter(
            Mandatory = $true, 
            Position = 0, 
            ParameterSetName = "ByGuid")]
        [String] $Guid,

        [String] $Destination,

        [switch] $DownloadAll
    )

    if ($Update) {
        $Guid = $Update.Guid | Select-Object -First 1
    }

    $Links = Get-UpdateLinks -Guid $Guid
    if (-not $Links) {
        Write-Warning "No valid download links found for GUID '$Guid'."
        return
    }

    $ProgressPreference = 'SilentlyContinue'
    $SuccessCount = 0
    $TotalCount = if ($DownloadAll) { $Links.Count } else { 1 }
    
    Write-Output "Found $($Links.Count) download links for GUID '$Guid'. $(if (-not $DownloadAll -and $Links.Count -gt 1) {"Using -DownloadAll to download all files."})"

    $LinksToProcess = if ($DownloadAll) { $Links } else { $Links | Select-Object -First 1 }

    foreach ($Link in $LinksToProcess) {
        $url = $Link.URL
        $name = $url.Split('/')[-1]
        $cleanname = $name.Split('_')[0]
        $extension = ".msu"
        $CleanOutFile = $cleanname + $extension

        $OutFile = Join-Path -Path $Destination -ChildPath $CleanOutFile
        
        if (Test-Path -Path $OutFile) {
            Write-Warning "File already exists: $CleanOutFile. Skipping download."
            continue
        }

        try {
            Write-Output "Downloading $CleanOutFile..."
            Set-TempSecurityProtocol
            Invoke-WebRequest -Uri $url -OutFile $OutFile -ErrorAction Stop
            Set-TempSecurityProtocol -ResetToDefault
            
            if (Test-Path -Path $OutFile) {
                Write-Output "Successfully downloaded file $CleanOutFile to $Destination"
                $SuccessCount++
            }
            else {
                Write-Warning "Downloading file $CleanOutFile failed."
            }
        }
        catch {
            Write-Warning "Error downloading $CleanOutFile : $_"
        }
    }

    Write-Output "Download complete: $SuccessCount of $TotalCount files downloaded successfully."
}