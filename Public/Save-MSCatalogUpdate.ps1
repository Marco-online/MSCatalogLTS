<#
    .SYNOPSIS
        This command is used to download update files from the https://www.catalog.update.microsoft.com website.

    .EXAMPLE
        Save-MSCatalogUpdate -Update $Updates -Destination ".\"
#>
function Save-MSCatalogUpdate {
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ParameterSetName = "ByObject"
        )]
        [Object] $Update,

        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipelineByPropertyName = "Guid",
            ParameterSetName = "ByGuid"
        )]
        [String] $Guid,

        [Parameter(
            Mandatory = $true,
            Position = 1,
            ParameterSetName = "ByObject"
        )]
        [Parameter(
            Mandatory = $true,
            Position = 1,
            ParameterSetName = "ByGuid"
        )]
        [String] $Destination
    )

    if ($Update) {
        $Guid = $Update.Guid | Select-Object -First 1
    }
    
    $Links = Get-UpdateLinks -Guid $Guid
    if ($Links.Count -eq 1) {
        Write-Verbose "Guid = $guid"

        $filename = $Links.Split('/')[-1]
        $cleanFilename = $filename.Split('_')[0]
        $extension = ".msu"
        $cleanFilenameWithExtension = $cleanFilename + $extension
        $OutFile = Join-Path -Path (Get-Item -Path $Destination) -ChildPath $cleanFilenameWithExtension

        Write-Verbose "$OutFile"
		$ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $Links -OutFile $OutFile
        }
   }

