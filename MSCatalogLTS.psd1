#
# Module manifest for module 'MSCatalogLTS'
#

@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'MSCatalogLTS.psm1'

    # Version number of this module.
    ModuleVersion = '2.1.0.0'

    # ID used to uniquely identify this module
    GUID = '721ac2a2-e4b6-4948-9c22-6ad2a52c0de6'

    # Author of this module
    Author = 'Marco-online'

    # Company or vendor of this module
    CompanyName = ''

    # Copyright statement for this module
    Copyright = '(c) 2026 Marco-online. All rights reserved.'

    # Description of the functionality provided by this module
    Description = 'MSCatalogLTS is a Long-term support module for searching and downloading Windows updates'

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '5.1.0.0'

    # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    DotNetFrameworkVersion = '4.5.1'

    # Format files (.ps1xml) to be loaded when importing this module
    FormatsToProcess = @(
        '.\Format\MSCatalogUpdate.Format.ps1xml'
    )

    # Functions to export from this module
    FunctionsToExport = @(
        'Get-MSCatalogUpdate',
        'Save-MSCatalogUpdate',
        'Save-MSCatalogOutput'
    )

    # Cmdlets to export from this module
    CmdletsToExport = @()

    # Variables to export from this module
    VariablesToExport = @()

    # Aliases to export from this module
    AliasesToExport = @()

    # Private data to pass to the module specified in RootModule/ModuleToProcess
    PrivateData = @{
        PSData = @{
            # Tags applied to this module for discoverability
            Tags = @(
                'Windows', 'Update', 'Catalog', 'Microsoft', 'WSUS', 
                'Patch', 'Security', 'WindowsUpdate', 'Edge', 'DotNet',
                'WindowsServer', 'Windows10', 'Windows11', 'Cumulative',
                'SecurityUpdate', 'Download', 'Automation', 'PSEdition_Desktop',
                'PSEdition_Core', 'Linux', 'Offline', 'Download',  'CrossPlatform'
            )

            # A URL to the license for this module
            LicenseUri = 'https://github.com/Marco-online/MSCatalogLTS/blob/main/LICENSE'

            # A URL to the main website for this project
            ProjectUri = 'https://github.com/Marco-online/MSCatalogLTS'

            # A URL to an icon representing this module
            # IconUri = ''

            # Release notes for this version
            ReleaseNotes = @'
Version 2.1.0.0 - February 2026

NEW FEATURES:
* Auto-strict search for Windows queries
* Edge channel filtering (-IsStable, -IsExtendedStable, -IsDev)
* Direct export with duplicate detection (-ExportJson, -ExportCsv, -ExportXml, -Append)
* Improved -Debug mode (no confirmation prompts)
* Smart update type detection (Servicing Stack, Dynamic Update)
* .NET Framework and Windows Server R2 filtering support

IMPROVEMENTS:
* -Descending defaults to true (most recent first)
* Upgraded HtmlAgilityPack to 1.12.4
* Enhanced error messages with examples
* Comprehensive debug logging
* Completely rewritten documentation

EXAMPLES:
  Get-MSCatalogUpdate -Search "Windows 11 24H2"
  Get-MSCatalogUpdate -Search "Edge x64" -IsStable
  Get-MSCatalogUpdate -Search "Windows Server 2022" -ExportJson "updates.json" -Append

Fully backward compatible with v1.x

More info: https://github.com/Marco-online/MSCatalogLTS
'@
        }
    }
}