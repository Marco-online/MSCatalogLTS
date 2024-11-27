<#
    .SYNOPSIS
        This command is used to download update files from the https://www.catalog.update.microsoft.com website.

    .EXAMPLE
        Save-MSCatalogUpdate -Update $Update -Destination ".\" -ShowDebug
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
        [String] $Destination,

        [Switch] $ShowDebug
    )

    if ($Update) {
        $Guid = $Update.Guid | Select-Object -First 1
    }
    
    $Links = Get-UpdateLinks -Guid $Guid
    if ($Links.Count -eq 1) {
         if ($ShowDebug) {
            Write-Host "DEBUG: Found GUID $($Guid)" -ForegroundColor yellow
            Write-Host "DEBUG: Download link $($Links)" -ForegroundColor yellow
        }
        $filename = $Links.Split('/')[-1]
        $cleanFilename = $filename.Split('_')[0]
        $extension = ".msu"
        $cleanFilenameWithExtension = $cleanFilename + $extension
        $OutFile = Join-Path -Path (Get-Item -Path $Destination) -ChildPath $cleanFilenameWithExtension

		$ProgressPreference = 'SilentlyContinue'
        if ($ShowDebug) {
             Write-Host "DEBUG: Download file $($cleanFilenameWithExtension) to $($Destination)" -ForegroundColor yellow
        }
        Invoke-WebRequest -Uri $Links -OutFile $OutFile

        if ($ShowDebug) {
        if (Test-Path -Path $OutFile) {
            Write-Host "DEBUG: File $cleanFilenameWithExtension successfully downloaded" -ForegroundColor yellow
        } else {
            Write-Warning "Downloading file $cleanFilenameWithExtension failed"
        }
        }
    }
   }

