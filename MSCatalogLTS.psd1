#
# Module manifest for module 'MSCatalogLTS'
#

@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'MSCatalogLTS.psm1'

    # Version number of this module.
    ModuleVersion = '2.1.0.2'

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
Version 2.1.0.2 - May 2026

**Bug fixes**
- Fix Get-MSCatalogUpdate -Search "<HWID GUID>" returning zero results — date-token regex no longer matches digit substrings inside GUIDs, KB numbers, or build versions
- Fix inverted Get-Module -ListAvailable check in Save-MSCatalogOutput (would attempt import only when ImportExcel was *not* installed)

**New: SupportUrl on update objects**
- New [string] $SupportUrl property on MSCatalogUpdate
- New private helper Get-UpdateSupportUrl (fetches from ScopedViewInline.aspx, parses suportUrlDiv)
- Save-MSCatalogOutput fetches SupportUrl per row when empty

**Save-MSCatalogOutput improvements**
- Excel column renamed Guid to UpdateID; new SupportUrl column
- Auto-creates worksheet if missing; rewrites sorted by LastUpdated
- Worksheet tabs sorted: numeric (01, 02, …) by value, non-numeric alphabetically at end
- Dedup checks both legacy Guid and new UpdateID columns
- Legacy Guid-column sheets are migrated to the new schema on next write

Fully backward compatible with v1.x

More info: https://github.com/Marco-online/MSCatalogLTS
'@
        }
    }
}