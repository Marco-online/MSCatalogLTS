function Get-MSCatalogUpdate {
    [CmdletBinding(DefaultParameterSetName = 'Search')]
    [OutputType([MSCatalogUpdate[]])]
    param (
        #region Parameters
        [Parameter(Mandatory = $false,
            HelpMessage = "Filter updates by architecture")]
        [ValidateSet("All", "x64", "x86", "arm64")]
        [string] $Architecture = "All",
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Sort in descending order")]
        [switch] $Descending,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Exclude .NET Framework updates")]
        [switch] $ExcludeFramework,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Filter updates from this date")]
        [DateTime] $FromDate,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Format for the results")]
        [ValidateSet("Default", "CSV", "JSON", "XML")]
        [string] $Format = "Default",
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Only show .NET Framework updates")]
        [switch] $GetFramework,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Search through all available pages")]
        [switch] $AllPages,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Include dynamic updates")]
        [switch] $IncludeDynamic,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Include file names in the results")]
        [switch] $IncludeFileNames,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Include preview updates")]
        [switch] $IncludePreview,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Filter updates from the last N days")]
        [int] $LastDays,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Filter updates with maximum size")]
        [double] $MaxSize,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Filter updates with minimum size")]
        [double] $MinSize,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'OS',
            HelpMessage = "Operating System to search updates for")]
        [ValidateSet("Windows 11", "Windows 10", "Windows Server")]
        [string] $OperatingSystem,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Select specific properties to display")]
        [string[]] $Properties,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Search',
            Position = 0,
            HelpMessage = "Search query for Microsoft Update Catalog")]
        [string] $Search,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Unit for size filtering (MB or GB)")]
        [ValidateSet("MB", "GB")]
        [string] $SizeUnit = "MB",
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Sort results by specified field")]
        [ValidateSet("Date", "Size", "Title", "Classification", "Product")]
        [string] $SortBy = "Date",
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Use strict search with exact phrase matching")]
        [switch] $Strict,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Filter updates until this date")]
        [DateTime] $ToDate,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Filter by update type")]
        [ValidateSet(
            "Security Updates", 
            "Updates", 
            "Critical Updates", 
            "Feature Packs", 
            "Service Packs", 
            "Tools", 
            "Update Rollups",
            "Cumulative Updates",
            "Security Quality Updates",
            "Driver Updates"
        )]
        [string[]] $UpdateType,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'OS',
            HelpMessage = "OS Version/Release (e.g., 22H2, 21H2, 23H2, 2019, 2016, 2012 R2)")]
        [string] $Version,

        [Parameter(Mandatory = $false,
            HelpMessage = "Include Microsoft Edge Stable channel updates")]
        [switch] $IsStable,

        [Parameter(Mandatory = $false,
            HelpMessage = "Include Microsoft Edge Extended Stable channel updates")]
        [switch] $IsExtendedStable,

        [Parameter(Mandatory = $false,
            HelpMessage = "Include Microsoft Edge Dev channel updates")]
        [switch] $IsDev,
        
        [Parameter(Mandatory = $false,
            HelpMessage = "Export results to JSON file")]
        [string] $ExportJson,

        [Parameter(Mandatory = $false,
            HelpMessage = "Export results to CSV file")]
        [string] $ExportCsv,

        [Parameter(Mandatory = $false,
            HelpMessage = "Export results to XML file")]
        [string] $ExportXml,

        [Parameter(Mandatory = $false,
            HelpMessage = "Append to existing export file instead of overwriting")]
        [switch] $Append
        #endregion Parameters
    )

    begin {
        #region Debug Preference Override
        if ($PSBoundParameters.ContainsKey('Debug') -and $PSBoundParameters['Debug']) {
            $DebugPreference = 'Continue'
            Write-Debug "Debug mode enabled - DebugPreference set to 'Continue'"
        }
        #endregion Debug Preference Override

        #region Parameter Validation
        if ($PSCmdlet.ParameterSetName -eq 'Search' -and $PSBoundParameters.ContainsKey('Version')) {
            $errorMessage = "The -Version parameter can only be used with -OperatingSystem parameter, not with -Search. " +
                            "Either use: -Search 'Windows 10 22H2' OR -OperatingSystem 'Windows 10' -Version '22H2'"
            if ($DebugPreference -ne 'SilentlyContinue') {
                Write-Error $errorMessage -ErrorAction Stop
            } else {
                Write-Host "Error: " -ForegroundColor Red -NoNewline
                Write-Host "Invalid parameter combination. Use -Search 'Windows 10 22H2' format or separate -OperatingSystem and -Version parameters."
                return
            }
        }
        #endregion Parameter Validation

        #region Initialization
        if (-not ('MSCatalogUpdate' -as [type])) {
            $classPath = Join-Path $PSScriptRoot '..\Classes\MSCatalogUpdate.Class.ps1'
            if (Test-Path $classPath) {
                . $classPath
            } else {
                throw "MSCatalogUpdate class file not found at: $classPath"
            }
        }

        $ProgressPreference = "SilentlyContinue"
        $Updates = @()
        $MaxResults = 1000
        $UseOSBehavior = $false
        $ShouldAbort = $false
        #endregion Initialization

        #region Smart Search Parser
        if ($PSCmdlet.ParameterSetName -eq 'Search' -and $Search) {
            # Apply default descending sort
            if (-not $PSBoundParameters.ContainsKey('Descending')) {
                $Descending = $true
                Write-Verbose "Applied default Descending=$Descending for Search parameter"
            }
            
            # Check for update type prefix in search (Servicing Stack, Dynamic Update, etc.)
            # BUT NOT .NET Framework - that's a product name, not an update type
            $updateTypePrefix = ""
            if ($Search -match '^(Servicing Stack|Dynamic Update|Security Update|Cumulative Update|Feature Update)\s+(.+)$' -and $Search -notmatch '\.NET') {
                $updateTypePrefix = $matches[1]
                $remainingSearch = $matches[2]
                Write-Debug "Detected update type prefix: '$updateTypePrefix', remaining search: '$remainingSearch'"
                
                $script:UpdateTypePrefix = $updateTypePrefix
                $Search = $remainingSearch
            } else {
                $script:UpdateTypePrefix = $null
            }
            
            # Special handling for .NET Framework searches
            $isDotNetFrameworkSearch = $Search -match '^\.NET Framework\s+(.+)$'
            if ($isDotNetFrameworkSearch) {
                $remainingAfterDotNet = $matches[1]
                Write-Debug "Detected .NET Framework search, remaining: '$remainingAfterDotNet'"
                
                # Check if there's a Windows version in the remaining part
                if ($remainingAfterDotNet -match '(Windows Server|Windows 10|Windows 11)\s+(?:Version\s+)?(\d{2}H[12]|\d{4})(\s+R2)?(?:\s+(x64|x86|arm64))?') {
                    $OperatingSystem = $matches[1]
                    $Version = $matches[2]
                    $R2Suffix = $matches[3]
                    $UseOSBehavior = $true
                    $script:IsDotNetFrameworkSearch = $true
                    
                    Write-Verbose "Parsed .NET Framework search: OS='$OperatingSystem', Version='$Version'"
                    Write-Debug ".NET Framework search with OS: OS='$($matches[1])', Version='$($matches[2])', R2='$($matches[3])', Arch='$($matches[4])'"
                    
                    # Handle R2 suffix
                    if ($R2Suffix) {
                        $script:IsR2Version = $true
                        $Version = "$Version R2"
                        Write-Debug "R2 version detected, updated Version to: $Version"
                    } else {
                        $script:IsR2Version = $false
                    }
                    
                    # Handle architecture
                    if ($matches[4]) {
                        if (-not $PSBoundParameters.ContainsKey('Architecture')) {
                            $Architecture = $matches[4]
                            Write-Verbose "Detected architecture from search string: $Architecture"
                            Write-Debug "Architecture parameter updated to: $Architecture"
                        } else {
                            Write-Debug "Architecture parameter was explicitly set to: $Architecture (ignoring detected: $($matches[4]))"
                        }
                    }
                    
                    $script:DotNetParsed = $true
                }
            } else {
                $script:IsDotNetFrameworkSearch = $false
                $script:DotNetParsed = $false
            }
            
            # Check if search contains OS name without version
            if (-not $script:DotNetParsed -and $Search -match '^(Windows Server|Windows 10|Windows 11)(?:\s+(x64|x86|arm64))?\s*$') {
                $detectedOS = $matches[1]
                Write-Host "Error: " -ForegroundColor Red -NoNewline
                Write-Host "Version is required when searching for $detectedOS updates."
                Write-Host ""
                Write-Host "Example usage:" -ForegroundColor Cyan
                
                switch ($detectedOS) {
                    "Windows 11" {
                        Write-Host "  Get-MSCatalogUpdate -Search 'Windows 11 23H2 x64'" -ForegroundColor Yellow
                        Write-Host "  Get-MSCatalogUpdate -Search 'Windows 11 24H2 x64'" -ForegroundColor Yellow
                    }
                    "Windows 10" {
                        Write-Host "  Get-MSCatalogUpdate -Search 'Windows 10 1809 x64'" -ForegroundColor Yellow
                        Write-Host "  Get-MSCatalogUpdate -Search 'Windows 10 22H2 x64'" -ForegroundColor Yellow
                    }
                    "Windows Server" {

                        Write-Host "  Get-MSCatalogUpdate -Search 'Windows Server 2012 R2 x64'" -ForegroundColor Yellow
                        Write-Host "  Get-MSCatalogUpdate -Search 'Windows Server 2025 x64'" -ForegroundColor Yellow
                    }
                }
                
                Write-Host ""
                $ShouldAbort = $true
                return
            }
            
            # Pattern matching: OS + Version + R2 (optional) + Architecture (optional)
            if (-not $script:DotNetParsed -and $Search -match '(Windows Server|Windows 10|Windows 11)\s+(?:Version\s+)?(\d{2}H[12]|\d{4})(\s+R2)?(?:\s+(x64|x86|arm64))?') {
                $OperatingSystem = $matches[1]
                $Version = $matches[2]
                $R2Suffix = $matches[3]
                $UseOSBehavior = $true
                
                Write-Verbose "Parsed Search: OperatingSystem='$OperatingSystem', Version='$Version', R2='$R2Suffix'"
                Write-Debug "Regex matches: OS='$($matches[1])', Version='$($matches[2])', R2='$($matches[3])', Arch='$($matches[4])'"
                
                # Handle R2 suffix
                if ($R2Suffix) {
                    $script:IsR2Version = $true
                    $Version = "$Version R2"
                    Write-Debug "R2 version detected, updated Version to: $Version"
                } else {
                    $script:IsR2Version = $false
                }
                
                # Handle architecture
                if ($matches[4]) {
                    if (-not $PSBoundParameters.ContainsKey('Architecture')) {
                        $Architecture = $matches[4]
                        Write-Verbose "Detected architecture from search string: $Architecture"
                        Write-Debug "Architecture parameter updated to: $Architecture"
                    } else {
                        Write-Debug "Architecture parameter was explicitly set to: $Architecture (ignoring detected: $($matches[4]))"
                    }
                }
            }
        }

        # OS parameter set
        if ($PSCmdlet.ParameterSetName -eq 'OS') {
            $UseOSBehavior = $true
            $script:IsR2Version = $false
            $script:UpdateTypePrefix = $null
            $script:IsDotNetFrameworkSearch = $false
            
            if (-not $PSBoundParameters.ContainsKey('Descending')) {
                $Descending = $true
                Write-Verbose "Applied default Descending=$Descending for OS parameter set"
            }
        }

        Write-Debug "After Smart Search Parser: Architecture='$Architecture', UseOSBehavior='$UseOSBehavior', IsR2='$script:IsR2Version', UpdateTypePrefix='$script:UpdateTypePrefix', IsDotNetFramework='$script:IsDotNetFrameworkSearch'"
        #endregion Smart Search Parser

        #region Query Building
        if (-not $ShouldAbort) {
            Write-Debug "Query Building: Architecture parameter = '$Architecture'"
            
            $searchQuery = if ($UseOSBehavior) {
                # Map Windows Server versions
                $mappedVersion = $Version
                if ($OperatingSystem -eq "Windows Server" -and $Version -notlike "*R2") {
                    $mappedVersion = switch ($Version) {
                        "2025" { "24H2" }
                        "2022" { "21H2" }
                        default { $Version }
                    }
                    Write-Debug "Windows Server version mapping: $Version -> $mappedVersion"
                }
                
                $script:VersionForFiltering = $mappedVersion
                
                # Special handling for .NET Framework searches
                if ($script:IsDotNetFrameworkSearch) {
                    $osPhrase = switch ($OperatingSystem) {
                        "Windows 10" { "Windows 10, version $mappedVersion" }
                        "Windows 11" { "Windows 11, version $mappedVersion" }
                        "Windows Server" { "Windows Server $mappedVersion" }
                        default { "$OperatingSystem $mappedVersion" }
                    }
                    
                    $archSuffix = if ($Architecture -ne "All") {
                        " for $Architecture"
                    } else {
                        ""
                    }
                    
                    $query = ".NET Framework $osPhrase$archSuffix"
                    Write-Debug ".NET Framework search query: '$query'"
                    $query
                } else {
                    # Regular OS query building
                    $osPhrase = switch ($OperatingSystem) {
                        "Windows 10" { "Windows 10 Version $mappedVersion" }
                        "Windows 11" { "Windows 11 Version $mappedVersion" }
                        "Windows Server" { 
                            if ($mappedVersion -match '^\d{2}H[12]$') {
                                "Microsoft Server Operating System version $mappedVersion"
                            } else {
                                "Windows Server $mappedVersion"
                            }
                        }
                        default { "$OperatingSystem $mappedVersion" }
                    }
                    
                    Write-Debug "OS Phrase: '$osPhrase'"
                    
                    # Build architecture suffix
                    $archSuffix = if ($Architecture -ne "All") {
                        $suffix = switch ($Architecture) {
                            "x64" { " for x64-based Systems" }
                            "x86" { " for x86-based Systems" }
                            "arm64" { " for ARM64-based Systems" }
                        }
                        Write-Debug "Architecture suffix: '$suffix'"
                        $suffix
                    } else {
                        Write-Debug "Architecture is 'All', no suffix added"
                        ""
                    }
                    
                    # Determine update prefix
                    $updatePrefix = ""
                    if ($UpdateType -and $UpdateType.Count -gt 0) {
                        if ($UpdateType -contains "Cumulative Updates" -or 
                            $UpdateType -contains "Updates" -or 
                            $UpdateType -contains "Security Updates") {
                            $updatePrefix = "Cumulative Update"
                        } 
                        elseif ($UpdateType -contains "Critical Updates" -or 
                                $UpdateType -contains "Feature Packs" -or 
                                $UpdateType -contains "Service Packs" -or 
                                $UpdateType -contains "Tools" -or 
                                $UpdateType -contains "Update Rollups" -or
                                $UpdateType -contains "Driver Updates") {
                            $updatePrefix = ""
                        }
                        else {
                            $updatePrefix = "Update"
                        }
                    } else {
                        if ($OperatingSystem -eq "Windows Server" -and $mappedVersion -match '^\d{4}') {
                            $updatePrefix = "Update"
                        } else {
                            $updatePrefix = "Cumulative Update"
                        }
                    }
                    
                    Write-Debug "Update prefix: '$updatePrefix'"
                    if ($script:UpdateTypePrefix) {
                        Write-Debug "UpdateTypePrefix '$script:UpdateTypePrefix' will be used for POST-FILTERING only"
                    }
                    
                    # Construct full query
                    if ($Version) {
                        if ($updatePrefix) {
                            "$updatePrefix for $osPhrase$archSuffix"
                        } else {
                            "$osPhrase$archSuffix"
                        }
                    } else {
                        if ($updatePrefix) {
                            "$updatePrefix for $OperatingSystem$archSuffix"
                        } else {
                            "$OperatingSystem$archSuffix"
                        }
                    }
                }
            } else {
                $script:VersionForFiltering = $null
                $script:IsR2Version = $false
                
                if (-not $Strict) {
                    $cleanSearch = $Search -replace '\s+for\s+(x64|x86|arm64).*$', '' `
                                            -replace '\s+\((x64|x86|arm64)\).*$', '' `
                                            -replace '\s+(x64|x86|arm64)\s*-.*$', ''
                    $cleanSearch
                } else {
                    $Search
                }
            }

            Write-Verbose "Search query: $searchQuery"
            Write-Debug "Full search query being sent to catalog: $searchQuery"
            Write-Debug "UseOSBehavior: $UseOSBehavior"
            Write-Debug "UpdateType specified: $($UpdateType -join ', ')"
            Write-Debug "VersionForFiltering: $script:VersionForFiltering"
            Write-Debug "IsR2Version: $script:IsR2Version"
            Write-Debug "UpdateTypePrefix for filtering: $script:UpdateTypePrefix"
            Write-Debug "IsDotNetFrameworkSearch: $script:IsDotNetFrameworkSearch"
            Write-Debug "Will use strict search: $(($PSCmdlet.ParameterSetName -eq 'OS' -and $Version) -or $Strict)"
        }
        #endregion Query Building
    }

    process {
        if ($ShouldAbort) {
            return
        }
        
        try {
            #region Search Preparation
            # Auto-enable strict search if "Windows" is in the search query
            $autoStrictEnabled = $searchQuery -match 'Windows'
            $isNetSearch = $searchQuery -match '^\.NET'
            $useStrictSearch = $false

            if (-not $isNetSearch) {
                $useStrictSearch = $Strict -or ($PSCmdlet.ParameterSetName -eq 'OS' -and $Version) -or $autoStrictEnabled
            }

            if ($isNetSearch -and $Strict) {
                Write-Debug "Strict search disabled for .NET queries (not supported by Microsoft Catalog)"
            }

            if ($autoStrictEnabled -and -not $Strict) {
                Write-Debug "Auto-enabled strict search (detected 'Windows' in query)"
            }

            $EncodedSearch = switch ($true) {
                $useStrictSearch { [uri]::EscapeDataString('"' + $searchQuery + '"') }
                $GetFramework { [uri]::EscapeDataString("*$searchQuery*") }
                default { [uri]::EscapeDataString($searchQuery) }
            }

            $Uri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$EncodedSearch"
            Write-Debug "Request URI: $Uri"
            Write-Debug "Using strict search: $useStrictSearch $(if ($autoStrictEnabled -and -not $Strict) { '(auto-enabled)' })"

            $Res = Invoke-CatalogRequest -Uri $Uri
            $Rows = $Res.Rows

            Write-Debug "Initial rows returned: $($Rows.Count)"
            #endregion Search Preparation

            #region Pagination
            if ($AllPages) {
                $PageCount = 0
                while ($Res.NextPage -and $PageCount -lt 39) {
                    $PageCount++
                    $PageUri = "$Uri&p=$PageCount"
                    Write-Debug "Fetching page $PageCount"
                    $Res = Invoke-CatalogRequest -Uri $PageUri
                    $Rows += $Res.Rows
                }
                Write-Debug "Total rows after pagination: $($Rows.Count)"
            }
            #endregion Pagination

            #region Base Filtering
            $Rows = $Rows.Where({
                $title = $_.SelectNodes("td")[1].InnerText.Trim()
                $classification = $_.SelectNodes("td")[3].InnerText.Trim()
                $include = $true
                
                Write-Debug "Evaluating: $title | Classification: $classification"
                
                # Basic exclusions
                if (-not $IncludeDynamic -and $title -like "*Dynamic*") { 
                    Write-Debug "Excluded (Dynamic): $title"
                    $include = $false 
                }
                if (-not $IncludePreview -and $title -like "*Preview*") { 
                    Write-Debug "Excluded (Preview): $title"
                    $include = $false 
                }
                
                # Framework filtering
                if ($GetFramework) {
                    if (-not ($title -like "*Framework*")) { 
                        Write-Debug "Excluded (Not Framework): $title"
                        $include = $false 
                    }
                } elseif ($ExcludeFramework) {
                    if ($title -like "*Framework*") { 
                        Write-Debug "Excluded (Framework): $title"
                        $include = $false 
                    }
                }

                # Edge channel filtering
                if ($IsStable -or $IsExtendedStable -or $IsDev) {
                    $isEdgeSearch = $title -match "Microsoft Edge"
                    
                    if ($isEdgeSearch) {
                        $includeByChannel = $false
                        
                        if ($IsStable -and $title -like "*Edge-Stable Channel*" -and $title -notlike "*Extended*") {
                            $includeByChannel = $true
                            Write-Debug "Included (Stable channel): $title"
                        }
                        if ($IsExtendedStable -and $title -like "*Extended Stable Channel*") {
                            $includeByChannel = $true
                            Write-Debug "Included (Extended Stable channel): $title"
                        }
                        if ($IsDev -and $title -like "*Edge-Dev Channel*") {
                            $includeByChannel = $true
                            Write-Debug "Included (Dev channel): $title"
                        }
                        
                        if (-not $includeByChannel) {
                            Write-Debug "Excluded (Edge channel filter - looking for: Stable=$IsStable, Extended=$IsExtendedStable, Dev=$IsDev): $title"
                            $include = $false
                        }
                    }
                }

                # Update type prefix filtering
                if ($script:UpdateTypePrefix) {
                    if (-not ($title -like "*$script:UpdateTypePrefix*")) {
                        Write-Debug "Excluded (Update type prefix '$script:UpdateTypePrefix' not found): $title"
                        $include = $false
                    }
                }
                
                # OS and Version filtering
                if ($UseOSBehavior) {
                    if ($OperatingSystem -eq "Windows Server") {
                        if (-not ($title -like "*Microsoft*Server*" -or 
                                  $title -like "*Server Operating System*" -or 
                                  $title -like "*Windows Server*")) { 
                            Write-Debug "Excluded (Not Windows Server): $title"
                            $include = $false 
                        }
                    } else {
                        if (-not ($title -like "*$OperatingSystem*")) { 
                            Write-Debug "Excluded (Not $OperatingSystem): $title"
                            $include = $false 
                        }
                    }
                    
                    $versionToCheck = if ($script:VersionForFiltering) { $script:VersionForFiltering } else { $Version }
                    if ($versionToCheck -and -not ($title -like "*$versionToCheck*")) { 
                        Write-Debug "Excluded (Not version $versionToCheck): $title"
                        $include = $false 
                    }
                    
                    # R2 filtering
                    if ($script:IsR2Version) {
                        $baseVersion = $versionToCheck -replace '\s+R2$', ''
                        if (($title -like "*$baseVersion*") -and -not ($title -like "*R2*")) {
                            Write-Debug "Excluded (R2 required but not found): $title"
                            $include = $false
                        }
                    } elseif ($versionToCheck -match '^\d{4}$' -and $OperatingSystem -eq "Windows Server") {
                        if ($title -like "*$versionToCheck R2*") {
                            Write-Debug "Excluded (Non-R2 required but R2 found): $title"
                            $include = $false
                        }
                    }
                }

                # UpdateType filtering
                if ($UpdateType) {
                    $hasMatchingType = $false
                    foreach ($type in $UpdateType) {
                        switch ($type) {
                            "Security Updates" { if ($classification -eq "Security Updates") { $hasMatchingType = $true } }
                            "Cumulative Updates" { if ($title -like "*Cumulative Update*") { $hasMatchingType = $true } }
                            "Critical Updates" { if ($classification -eq "Critical Updates") { $hasMatchingType = $true } }
                            "Updates" { if ($classification -eq "Updates") { $hasMatchingType = $true } }
                            "Feature Packs" { if ($classification -eq "Feature Packs") { $hasMatchingType = $true } }
                            "Service Packs" { if ($classification -eq "Service Packs") { $hasMatchingType = $true } }
                            "Tools" { if ($classification -eq "Tools") { $hasMatchingType = $true } }
                            "Update Rollups" { if ($classification -eq "Update Rollups") { $hasMatchingType = $true } }
                            "Security Quality Updates" { 
                                if (($classification -eq "Security Updates") -and ($title -like "*Quality Update*")) { 
                                    $hasMatchingType = $true 
                                } 
                            }
                            "Driver Updates" { if ($title -like "*Driver*") { $hasMatchingType = $true } }
                            default { if ($title -like "*$type*") { $hasMatchingType = $true } }
                        }
                        if ($hasMatchingType) { break }
                    }
                    if (-not $hasMatchingType) { 
                        Write-Debug "Excluded (UpdateType mismatch - looking for: $($UpdateType -join ', ')): $title"
                        $include = $false 
                    }
                }
                
                if ($include) {
                    Write-Debug "Included: $title"
                }
                
                $include
            })

            Write-Debug "Rows after base filtering: $($Rows.Count)"
            #endregion Base Filtering

            #region Architecture Filtering
            $skipArchFilter = ($PSCmdlet.ParameterSetName -eq 'OS' -and $Architecture -ne "All" -and $Version) -or 
                              ($UseOSBehavior -and $Architecture -ne "All" -and $Version)

            if ($Architecture -ne "All" -and -not $skipArchFilter) {
                $preFilterCount = $Rows.Count
                $Rows = $Rows.Where({
                    $title = $_.SelectNodes("td")[1].InnerText.Trim()
                    $match = switch ($Architecture) {
                        "x64" { $title -match "x64|64.?bit|64.?based" -and -not ($title -match "x86|32.?bit|arm64") }
                        "x86" { $title -match "x86|32.?bit|32.?based" -and -not ($title -match "64.?bit|arm64") }
                        "arm64" { $title -match "arm64|ARM.?based" }
                    }
                    if (-not $match) {
                        Write-Debug "Excluded (Architecture $Architecture): $title"
                    }
                    $match
                })
                Write-Debug "Rows after architecture filtering ($Architecture): $($Rows.Count) (was: $preFilterCount)"
            } elseif ($skipArchFilter) {
                Write-Debug "Skipping architecture filter (already in search query)"
            }
            #endregion Architecture Filtering

            #region Create Update Objects
            $Updates = $Rows.Where({ $_.Id -ne "headerRow" }).ForEach({
                try {
                    [MSCatalogUpdate]::new($_, $IncludeFileNames)
                } catch {
                    if ($DebugPreference -ne 'SilentlyContinue') {
                        Write-Warning "Failed to process update: $($_.Exception.Message)"
                    }
                    $null
                }
            }) | Where-Object { $null -ne $_ }
            
            Write-Debug "Updates created: $($Updates.Count)"
            #endregion Create Update Objects

            #region Apply Filters
            if ($FromDate) { 
                $preCount = $Updates.Count
                $Updates = $Updates.Where({ $_.LastUpdated -ge $FromDate }) 
                Write-Debug "After FromDate filter: $($Updates.Count) (was: $preCount)"
            }
            if ($ToDate) { 
                $preCount = $Updates.Count
                $Updates = $Updates.Where({ $_.LastUpdated -le $ToDate }) 
                Write-Debug "After ToDate filter: $($Updates.Count) (was: $preCount)"
            }
            if ($LastDays) {
                $CutoffDate = (Get-Date).AddDays(-$LastDays)
                $preCount = $Updates.Count
                $Updates = $Updates.Where({ $_.LastUpdated -ge $CutoffDate })
                Write-Debug "After LastDays filter ($LastDays days, cutoff: $CutoffDate): $($Updates.Count) (was: $preCount)"
            }

            if ($MinSize -or $MaxSize) {
                $Multiplier = if ($SizeUnit -eq "GB") { 1024 } else { 1 }
                $preCount = $Updates.Count
                $Updates = $Updates.Where({
                    $size = [double]($_.Size -replace ' MB$','')
                    $meetsMin = -not $MinSize -or $size -ge ($MinSize * $Multiplier)
                    $meetsMax = -not $MaxSize -or $size -le ($MaxSize * $Multiplier)
                    $meetsMin -and $meetsMax
                })
                Write-Debug "After size filter: $($Updates.Count) (was: $preCount)"
            }
            #endregion Apply Filters

            #region Sorting and Output
            # Apply sorting
            $Updates = switch ($SortBy) {
                "Date" { $Updates | Sort-Object LastUpdated -Descending:$Descending }
                "Size" { $Updates | Sort-Object { [double]($_.Size -replace ' MB$','') } -Descending:$Descending }
                "Title" { $Updates | Sort-Object Title -Descending:$Descending }
                "Classification" { $Updates | Sort-Object Classification -Descending:$Descending }
                "Product" { $Updates | Sort-Object Products -Descending:$Descending }
                default { $Updates }
            }

            # Handle exports
            $exportPerformed = $false

            if ($ExportJson) {
                try {
                    $jsonContent = if ($Properties) { 
                        $Updates | Select-Object $Properties | ConvertTo-Json -Depth 10
                    } else { 
                        $Updates | ConvertTo-Json -Depth 10
                    }
                    
                    if ($Append -and (Test-Path $ExportJson)) {
                        $existingJson = Get-Content $ExportJson -Raw | ConvertFrom-Json
                        $newJson = $jsonContent | ConvertFrom-Json
                        
                        $existingArray = @($existingJson)
                        $newArray = @($newJson)
                        
                        $existingGuids = @{}
                        foreach ($item in $existingArray) {
                            if ($item.Guid) {
                                $existingGuids[$item.Guid] = $true
                            }
                        }
                        
                        $uniqueNewItems = @()
                        $duplicateCount = 0
                        foreach ($item in $newArray) {
                            if ($item.Guid -and $existingGuids.ContainsKey($item.Guid)) {
                                $duplicateCount++
                                Write-Debug "Skipping duplicate: $($item.Title) (GUID: $($item.Guid))"
                            } else {
                                $uniqueNewItems += $item
                                if ($item.Guid) {
                                    $existingGuids[$item.Guid] = $true
                                }
                            }
                        }
                        
                        $combined = $existingArray + $uniqueNewItems
                        $combined | ConvertTo-Json -Depth 10 | Out-File $ExportJson -Encoding UTF8
                        
                        $message = "Appended $($uniqueNewItems.Count) updates to JSON file: $ExportJson"
                        if ($duplicateCount -gt 0) {
                            $message += " (skipped $duplicateCount duplicate$(if($duplicateCount -gt 1){'s'}))"
                        }
                        Write-Verbose $message
                        if ($duplicateCount -gt 0) {
                            Write-Host $message -ForegroundColor Yellow
                        }
                    } else {
                        $jsonContent | Out-File $ExportJson -Encoding UTF8
                        Write-Verbose "Exported $($Updates.Count) updates to JSON file: $ExportJson"
                    }
                    $exportPerformed = $true
                } catch {
                    Write-Warning "Failed to export to JSON: $($_.Exception.Message)"
                }
            }

            if ($ExportCsv) {
                try {
                    if ($Append -and (Test-Path $ExportCsv)) {
                        $existingCsv = Import-Csv $ExportCsv
                        
                        $existingGuids = @{}
                        foreach ($item in $existingCsv) {
                            if ($item.Guid) {
                                $existingGuids[$item.Guid] = $true
                            }
                        }
                        
                        $uniqueNewItems = @()
                        $duplicateCount = 0
                        foreach ($item in $Updates) {
                            if ($item.Guid -and $existingGuids.ContainsKey($item.Guid)) {
                                $duplicateCount++
                                Write-Debug "Skipping duplicate: $($item.Title) (GUID: $($item.Guid))"
                            } else {
                                $uniqueNewItems += $item
                            }
                        }
                        
                        if ($uniqueNewItems.Count -gt 0) {
                            if ($Properties) {
                                $uniqueNewItems | Select-Object $Properties | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -Append
                            } else {
                                $uniqueNewItems | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -Append
                            }
                        }
                        
                        $message = "Appended $($uniqueNewItems.Count) updates to CSV file: $ExportCsv"
                        if ($duplicateCount -gt 0) {
                            $message += " (skipped $duplicateCount duplicate$(if($duplicateCount -gt 1){'s'}))"
                        }
                        Write-Verbose $message
                        if ($duplicateCount -gt 0) {
                            Write-Host $message -ForegroundColor Yellow
                        }
                    } else {
                        if ($Properties) {
                            $Updates | Select-Object $Properties | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
                        } else {
                            $Updates | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
                        }
                        Write-Verbose "Exported $($Updates.Count) updates to CSV file: $ExportCsv"
                    }
                    
                    $exportPerformed = $true
                } catch {
                    Write-Warning "Failed to export to CSV: $($_.Exception.Message)"
                }
            }

            if ($ExportXml) {
                try {
                    if ($Append -and (Test-Path $ExportXml)) {
                        [xml]$existingXml = Get-Content $ExportXml
                        $existingObjects = $existingXml.Objects.Object
                        
                        $existingGuids = @{}
                        foreach ($obj in $existingObjects) {
                            $guidProp = $obj.Property | Where-Object { $_.Name -eq 'Guid' }
                            if ($guidProp) {
                                $existingGuids[$guidProp.'#text'] = $true
                            }
                        }
                        
                        $uniqueNewItems = @()
                        $duplicateCount = 0
                        foreach ($item in $Updates) {
                            if ($item.Guid -and $existingGuids.ContainsKey($item.Guid)) {
                                $duplicateCount++
                                Write-Debug "Skipping duplicate: $($item.Title) (GUID: $($item.Guid))"
                            } else {
                                $uniqueNewItems += $item
                            }
                        }
                        
                        if ($uniqueNewItems.Count -gt 0) {
                            $xmlContent = if ($Properties) { 
                                $uniqueNewItems | Select-Object $Properties | ConvertTo-Xml -As String -Depth 10
                            } else { 
                                $uniqueNewItems | ConvertTo-Xml -As String -Depth 10
                            }
                            Add-Content -Path $ExportXml -Value $xmlContent -Encoding UTF8
                        }
                        
                        $message = "Appended $($uniqueNewItems.Count) updates to XML file: $ExportXml"
                        if ($duplicateCount -gt 0) {
                            $message += " (skipped $duplicateCount duplicate$(if($duplicateCount -gt 1){'s'}))"
                        }
                        Write-Verbose $message
                        if ($duplicateCount -gt 0) {
                            Write-Host $message -ForegroundColor Yellow
                        }
                    } else {
                        $xmlContent = if ($Properties) { 
                            $Updates | Select-Object $Properties | ConvertTo-Xml -As String -Depth 10
                        } else { 
                            $Updates | ConvertTo-Xml -As String -Depth 10
                        }
                        $xmlContent | Out-File $ExportXml -Encoding UTF8
                        Write-Verbose "Exported $($Updates.Count) updates to XML file: $ExportXml"
                    }
                    $exportPerformed = $true
                } catch {
                    Write-Warning "Failed to export to XML: $($_.Exception.Message)"
                }
            }

            # Display result summary
            $IsUpdate = ($MyInvocation.Line -match '^\s*\$update\s*=')
            $IsPiped = ($PSCmdlet.MyInvocation.PipelineLength -gt 1)

            if (-not $IsUpdate -and -not $IsPiped) {
                Write-Host "`nSearch completed for: $searchQuery"
                Write-Host "Found $($Updates.Count) updates"
                
                if ($exportPerformed) {
                    Write-Host ""
                    if ($ExportJson) { Write-Host "Exported to JSON: $ExportJson" -ForegroundColor Green }
                    if ($ExportCsv) { Write-Host "Exported to CSV: $ExportCsv" -ForegroundColor Green }
                    if ($ExportXml) { Write-Host "Exported to XML: $ExportXml" -ForegroundColor Green }
                }
                
                if ($Updates.Count -eq 0) {
                    Write-Host "`nNo updates found matching the criteria." -ForegroundColor Yellow
                    if ($DebugPreference -ne 'SilentlyContinue') {
                        Write-Host "Try:" -ForegroundColor Cyan
                        Write-Host "  - Removing or adjusting filters (-UpdateType, -Architecture, -LastDays)" -ForegroundColor Cyan
                        Write-Host "  - Using -AllPages to search more results" -ForegroundColor Cyan
                        Write-Host "  - Using -Verbose or -Debug for more information" -ForegroundColor Cyan
                    }
                }
            }

            if ($Updates.Count -ge $MaxResults) {
                Write-Warning "Result limit of $MaxResults reached. Please refine your search criteria."
            }

            # Return results to pipeline
            if (-not $exportPerformed -or $IsUpdate -or $IsPiped) {
                switch ($Format) {
                    "Default" { 
                        if ($Properties) { $Updates | Select-Object $Properties }
                        else { $Updates }
                    }
                    "CSV" { 
                        if ($Properties) { $Updates | Select-Object $Properties | ConvertTo-Csv -NoTypeInformation }
                        else { $Updates | ConvertTo-Csv -NoTypeInformation }
                    }
                    "JSON" { 
                        if ($Properties) { $Updates | Select-Object $Properties | ConvertTo-Json }
                        else { $Updates | ConvertTo-Json }
                    }
                    "XML" { 
                        if ($Properties) { $Updates | Select-Object $Properties | ConvertTo-Xml -As String }
                        else { $Updates | ConvertTo-Xml -As String }
                    }
                }
            }
            #endregion Sorting and Output
        }
        catch {
            if ($DebugPreference -ne 'SilentlyContinue') {
                Write-Warning "Error processing search request: $($_.Exception.Message)"
                Write-Debug "Full error: $($_.Exception | Format-List * -Force | Out-String)"
            }
        }
    }

    end {
        $ProgressPreference = "Continue"
    }
}