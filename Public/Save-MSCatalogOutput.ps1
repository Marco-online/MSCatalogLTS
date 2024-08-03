<#
    .SYNOPSIS
        Save output from Get-MSCatalogUpdate to csv file.

    .EXAMPLE
        Save-MSCatalogOutput -Update $update
#>

function Save-MSCatalogOutput {
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ParameterSetName = "ByObject"
        )]
        [Object] $Update
    )

    if ($Update) {
        $data += [PSCustomObject]@{
            Title      = $Update.Title
            Guid       = $Update.Guid
            LastUpdated = $Update.LastUpdated
        }
    }
    $date = Get-Date -Format "yyyyMMdd"
    $filePath = Join-Path -Path "$env:WINDIR\Temp" -ChildPath "WinPatches$date.csv"
    $data | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
}
