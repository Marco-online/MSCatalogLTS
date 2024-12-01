function Save-MSCatalogUpdate {
    param (
        [Parameter(
            Position = 0, 
            ParameterSetName = "ByObject")]
        [Object] $Update,

        [Parameter(
            Mandatory = $true, 
            Position = 0, 
            ParameterSetName = "ByGuid")]
        [String] $Guid,

        [String] $Destination
    )

    if ($Update) {
        $Guid = $Update.Guid | Select-Object -First 1
    }

    $Links = Get-UpdateLinks -Guid $Guid
    if ($Links.Count -eq 1) {

        $name = $Links.Split('/')[-1]
        $cleanname = $name.Split('_')[0]
        $extension = ".msu"
        $CleanOutFile = $cleanname + $extension

        $OutFile = Join-Path -Path $Destination -ChildPath $CleanOutFile
        $ProgressPreference = 'SilentlyContinue'

        if (Test-Path -Path $OutFile) {
            Write-Warning "File already exists: $CleanOutFile. Skipping download."
            return
        } else {
            Set-TempSecurityProtocol
            Invoke-WebRequest -Uri $Links -OutFile $OutFile -ErrorAction Stop
            Set-TempSecurityProtocol -ResetToDefault
        }

        if (Test-Path -Path $OutFile) {
            Write-Output "Successfully downloaded file $CleanOutFile to $Destination"
        } else {
            Write-Warning "Downloading file $CleanOutFile failed."
        }
    } else {
        Write-Warning "No valid download links found for GUID '$Guid'."
    }
}