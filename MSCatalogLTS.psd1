#
# Module manifest for module 'MSCatalogLTS'
#

@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'MSCatalogLTS.psm1'

    # Version number of this module.
    ModuleVersion = '2.1.0.1'

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
Version 2.1.0.1 - February 2026

NEW FEATURES:
* Added -Date Option Filter updates of specific date

IMPROVEMENTS:
* Fixed -GetFramework now works correctly
* .NET Framework is added to the catalog query (OS and literal flows) instead of only client-side filtering/wildcards
* Parser enhancement: smarter OS vs literal-search detection
* Only trigger OS-behavior when appropriate; update-type phrases are treated as literal search text.
* Query construction: refined for OS/version/arch and .NET Framework searches

Fully backward compatible with v1.x

More info: https://github.com/Marco-online/MSCatalogLTS
'@
        }
    }
}