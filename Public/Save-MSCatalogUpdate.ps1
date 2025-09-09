function Save-MSCatalogUpdate {
    param (
        [Parameter(
            Position = 0, 
            ParameterSetName = "ByObject",
            ValueFromPipeline = $true)]
        [Object] $Update,

        [Parameter(
            Mandatory = $true, 
            Position = 0, 
            ParameterSetName = "ByGuid")]
        [String[]] $Guid,

        [Parameter(Position = 1)]
        [String] $Destination = (Get-Location).Path,

        [switch] $DownloadAll
    )
    begin {
        # Check if destination directory exists and create it if needed
        if (-not (Test-Path -Path $Destination -PathType Container)) {
            try {
                New-Item -Path $Destination -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Write-Output "Created destination directory: $Destination"
            }
            catch {
                Write-Error "Failed to create destination directory '$Destination': $_" -ErrorAction Stop
            }
        }
    }
    process {
        $GuidsToProcess = if ($PSCmdlet.ParameterSetName -eq "ByObject") {
            $Update.Guid
        }
        else {
            $Guid
        }

        $ProgressPreference = 'SilentlyContinue'
        foreach ($GuidItem in $GuidsToProcess)
        {
            $Links = Get-MSCatalogUpdateLink -Guid $GuidItem
            if (-not $Links) {
                Write-Warning "No valid download links found for GUID '$GuidItem'."
                continue
            }

            $SuccessCount = 0
            $TotalCount = if ($DownloadAll) { $Links.Count } else { 1 }
    
            Write-Output "Found $($Links.Count) download links for GUID '$GuidItem'. $(if (-not $DownloadAll -and $Links.Count -gt 1) {"Only downloading the first file. Use -DownloadAll to download all files."})"

            $LinksToProcess = if ($DownloadAll) { $Links } else { $Links | Select-Object -First 1 }

            foreach ($Link in $LinksToProcess) {
                $url = $Link.URL
                $name = $url.Split('/')[-1]
                $cleanname = $name.Split('_')[0]
        
                # Determine extension based on URL or use .msu as default
                $extension = if ($url -match '\.(cab|exe|msi|msp|msu)$') {
                    ".$($matches[1])"
                } else {
                    ".msu"
                }
        
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
    }
}