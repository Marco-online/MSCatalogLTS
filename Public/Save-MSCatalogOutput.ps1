function Save-MSCatalogOutput {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [Object] $Update,

        [Parameter(Mandatory = $true)]
        [string] $Destination,

        [string] $WorksheetName = "Updates"
    )

    if (-not (Get-Module -Name ImportExcel)) {
        try { Import-Module ImportExcel -ErrorAction Stop }
        catch { Write-Warning "Unable to import ImportExcel"; return }
    }

    if ($Update.Count -gt 1) {
        $Update = $Update | Select-Object -First 1
    }

    # fetch SupportUrl if not already populated. The catalog detail page requires an extra HTTP request 
    if (-not $Update.SupportUrl) {
        $Update.SupportUrl = Get-UpdateSupportUrl -Guid $Update.Guid
    }

    # fetch SHA1 from the first download link.
    if (-not $Update.SHA1) {
        $FirstLink = Get-UpdateLinks -Guid $Update.Guid | Select-Object -First 1
        if ($FirstLink -and $FirstLink.SHA1) {
            $Update.SHA1 = [BitConverter]::ToString([Convert]::FromBase64String($FirstLink.SHA1)).Replace('-','')
        }
    }

    $data = [PSCustomObject]@{
        Title          = $Update.Title
        Products       = $Update.Products
        Classification = $Update.Classification
        LastUpdated    = $Update.LastUpdated.ToString('yyyy/MM/dd')
        UpdateID       = $Update.Guid
        SupportUrl     = $Update.SupportUrl
        SHA1           = $Update.SHA1
    }

    $filePath = $Destination
    $tableName = "Table_$WorksheetName"

    # Create workbook if missing
    if (-not (Test-Path $filePath)) {
        $data | Export-Excel -Path $filePath -WorksheetName $WorksheetName `
            -AutoSize -TableStyle Light1 -TableName $tableName -KillExcel
        return
    }

    # Ensure worksheet exists
    $sheetInfo = Get-ExcelSheetInfo -Path $filePath -ErrorAction SilentlyContinue
    if ($sheetInfo.Name -notcontains $WorksheetName) {
        $data | Export-Excel -Path $filePath -WorksheetName $WorksheetName `
            -AutoSize -TableStyle Light1 -TableName $tableName -KillExcel
    } else {
        # Prevent duplicates - check both UpdateID (new) and Guid (legacy) columns
        $existingData = Import-Excel -Path $filePath -WorksheetName $WorksheetName -DataOnly
        $existingIds = @($existingData.UpdateID) + @($existingData.Guid)
        if ($existingIds -contains $Update.Guid) { return }

        # Migrate legacy "Guid" rows to new schema (UpdateID + SupportUrl + SHA1)
        $existingData = $existingData | ForEach-Object {
            [PSCustomObject]@{
                Title          = $_.Title
                Products       = $_.Products
                Classification = $_.Classification
                LastUpdated    = $_.LastUpdated
                UpdateID       = if ($_.UpdateID) { $_.UpdateID } else { $_.Guid }
                SupportUrl     = $_.SupportUrl
                SHA1           = $_.SHA1
            }
        }

        # Merge new row, sort, and write once
        $allData = @($existingData) + @($data) | Sort-Object LastUpdated

        $allData | Export-Excel -Path $filePath -WorksheetName $WorksheetName `
            -AutoSize -TableStyle Light1 -TableName $tableName -ClearSheet -KillExcel
    }
        # Sort worksheet tabs numerically
        $excel = Open-ExcelPackage -Path $filePath
        $workbook = $excel.Workbook

        # Numeric sheets (01, 02, ...) sort by value; non-numeric sheets land at the end alphabetically
        $sortedNames = $workbook.Worksheets.Name |
            Sort-Object { if ($_ -match '^\d+$') { [int]$_ } else { [int]::MaxValue } }, { $_ }

        for ($i = $sortedNames.Count - 1; $i -ge 0; $i--) {
        try {
            $workbook.Worksheets.MoveToStart($sortedNames[$i])
            } catch {}
        }

        Close-ExcelPackage -ExcelPackage $excel -SaveAs $filePath
}
