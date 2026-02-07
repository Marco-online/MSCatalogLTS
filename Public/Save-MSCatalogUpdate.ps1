function Save-MSCatalogUpdate {
    param (
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [MSCatalogUpdate] $Update,

        [Parameter(Position = 1)]
        [String[]] $Guid,

        [Parameter(Position = 2)]
        [String] $Destination = (Get-Location).Path,

        [switch] $Force,

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
        # Fixes #21, Fixes #19
        # Drop empty entries 
        $GuidsToProcess = $GuidsToProcess | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        if (-not $GuidsToProcess -or $GuidsToProcess.Count -eq 0) {
             # Filter out updates with valid GUIDs
            $ValidUpdates = $AllUpdates | Where-Object { $_.Guid -and $_.Guid -ne '' }

            if ($ValidUpdates -and $ValidUpdates.Count -gt 0) {
                # Sort by Title to find the latest update
                $GuidsToProcess = $ValidUpdates |
                    Sort-Object -Property Title -Descending |
                    ForEach-Object { $_.Guid } |
                    Select-Object -First 1 |
                    Select-Object -Unique
                Write-Output "Selected $($GuidsToProcess.Count) update(s) from pipeline."
            } else {
                Write-Warning "No valid update found with a GUID."
                return
            }
        }

        foreach ($GuidItem in $GuidsToProcess) {
            if ([string]::IsNullOrWhiteSpace($GuidItem)) {
                Write-Warning "Skipped empty GUID."
                continue
            }

            # Get update links ONCE
            $Links = Get-UpdateLinks -Guid $GuidItem
            
            # Filter out invalid links (null, empty, or missing URL)
            if ($Links) {
                $Links = @($Links | Where-Object { $_ -and $_.URL -and $_.URL.Trim() -ne "" })
                
                # Remove duplicates
                $UniqueLinks = @()
                $SeenUrls = @{}
                foreach ($Link in $Links) {
                    if (-not $SeenUrls.ContainsKey($Link.URL)) {
                        $UniqueLinks += $Link
                        $SeenUrls[$Link.URL] = $true
                    }
                }
                $Links = $UniqueLinks
            }

            # Check if link is valid
            if (-not $Links -or $Links.Count -eq 0) {
                Write-Warning "No valid download links found for GUID '$GuidItem'."
                continue
            }

            $TotalCount = if ($DownloadAll) { $Links.Count } else { 1 }           
            Write-Output "Found $($Links.Count) download link$(if ($Links.Count -ne 1) {'s'}) for GUID '$GuidItem'. $(if (-not $DownloadAll -and $Links.Count -gt 1) {"Only downloading the first file. Use -DownloadAll to download all files."})"

            # Select links to process
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

                # Force overwrite check / Fixes#22
                if (Test-Path -Path $OutFile) {
                    if ($Force) {
                        # Do nothing — allow download to overwrite
                    }
                    else {
                        Write-Warning "File already exists: $CleanOutFile Skipping download."
                        continue
                    }
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
            Write-Output "Download complete: $SuccessCount of $TotalCount file$(if ($TotalCount -ne 1) {'s'}) downloaded successfully."
        }
    }
}