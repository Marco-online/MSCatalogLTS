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
        $existingData = Import-Excel -Path $filePath -WorksheetName $WorksheetName
        if ($existingData.Guid -contains $Update.Guid) {
        return
    }
        $data | Export-Excel -Path $filePath -WorksheetName $WorksheetName -Append -AutoSize -TableStyle Light1
    } else {
        $data | Export-Excel -Path $filePath -WorksheetName $WorksheetName -AutoSize -TableStyle Light1
    }
}