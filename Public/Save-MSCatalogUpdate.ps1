function Save-MSCatalogUpdate {
    param (
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [MSCatalogUpdate] $Update,

        [Parameter(Position = 1)]
        [String[]] $Guid,

        [Parameter(Position = 2)]
        [String] $Destination = (Get-Location).Path,

        [switch] $DownloadAll
    )

    begin {
        $ProgressPreference = 'SilentlyContinue'
        $GuidsToProcess = @()
        $AllUpdates = @()

        # Check if destination directory exists and create it if needed
        if (-not (Test-Path -Path $Destination -PathType Container)) {
            try {
                New-Item -Path $Destination -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Write-Output "Created destination directory: $Destination"
            }
            catch {
                Write-Error "Failed to create destination directory '$Destination': $_"
                return
            }
        }
    }

    process {
        if ($Update -and $Update.Guid) {
            $AllUpdates += $Update
        }
        if ($Guid) {
            $GuidsToProcess += $Guid
        }
    }

    end {
        # Filter out updates with valid GUIDs
        $ValidUpdates = $AllUpdates | Where-Object { $_.Guid -and $_.Guid -ne '' }

        # Sort by Title to find the latest update
        $LatestUpdate = $ValidUpdates |
            Sort-Object -Property Title -Descending |
            Select-Object -First 1

        if ($LatestUpdate -and $LatestUpdate.Guid) {
            Write-Output "Selected latest update: $($LatestUpdate.Title)"
            $GuidsToProcess = @($LatestUpdate.Guid)
        } else {
            Write-Warning "No valid update found with a GUID."
            return
        }

        foreach ($GuidItem in $GuidsToProcess) {
            if ([string]::IsNullOrWhiteSpace($GuidItem)) {
                Write-Warning "Skipped empty GUID."
                continue
            }

            $Links = Get-UpdateLinks -Guid $GuidItem
            if (-not $Links) {
                Write-Warning "No valid download links found for GUID '$GuidItem'."
                continue
            }

            $TotalCount = if ($DownloadAll) { $Links.Count } else { 1 }
            Write-Output "Found $($Links.Count) download links for GUID '$GuidItem'. $(if (-not $DownloadAll -and $Links.Count -gt 1) {"Only downloading the first file. Use -DownloadAll to download all files."})"

            $LinksToProcess = if ($DownloadAll) { $Links } else { $Links | Select-Object -First 1 }
            $SuccessCount = 0

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
                    } else {
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
