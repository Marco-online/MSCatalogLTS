<#
    .SYNOPSIS
        Save output from Get-MSCatalogUpdate to csv file.
    .EXAMPLE
        Save-MSCatalogOutput -Update $update -Destination ".\output.csv"
#>
function Save-MSCatalogOutput {
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ParameterSetName = "ByObject"
        )]
        [Object] $Update,
        [Parameter(Mandatory = $true)]
        [string] $Destination
    )

    if ($Update.Count -gt 1) {
        $Update = $Update | Select-Object -First 1
    }

    $filePath = $Destination
    $append = Test-Path -Path $filePath

    $data = [PSCustomObject]@{
        Title          = $Update.Title
        Products       = $Update.Products
        Classification = $Update.Classification
        LastUpdated    = $Update.LastUpdated.ToString('yyyy/MM/dd')
        Guid           = $Update.Guid
    }
    $data | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Append:$append
}
