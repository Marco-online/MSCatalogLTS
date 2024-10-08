﻿<#
    .SYNOPSIS
        Save output from Get-MSCatalogUpdate to csv file.
    .EXAMPLE
        Save-MSCatalogOutput -Update $update -WorksheetName "08_2024_Updates" -Destination "C:\Temp\2024_Updates.xlsx"
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
        [string] $Destination,

        [string] $WorksheetName = "Updates"
    )

    if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
        try {
            Import-Module ImportExcel -ErrorAction Stop
        }
        catch {
            Write-Warning "Unable to Import the Excel Module"
            return
        }
    }

    if ($Update.Count -gt 1) {
        $Update = $Update | Select-Object -First 1
    }

    $data = [PSCustomObject]@{
        Title          = $Update.Title
        Products       = $Update.Products
        Classification = $Update.Classification
        LastUpdated    = $Update.LastUpdated.ToString('yyyy/MM/dd')
        Guid           = $Update.Guid
    }

    $filePath = $Destination
    if (Test-Path -Path $filePath) {
        $data | Export-Excel -Path $filePath -WorksheetName $WorksheetName -Append -AutoSize -TableStyle Light1
    } else {
        $data | Export-Excel -Path $filePath -WorksheetName $WorksheetName -AutoSize -TableStyle Light1
    }
}
