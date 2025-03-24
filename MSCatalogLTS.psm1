try {
    if (!([System.Management.Automation.PSTypeName]'HtmlAgilityPack.HtmlDocument').Type) {
        if ($PSVersionTable.PSEdition -eq "Desktop") {
            Add-Type -Path "$PSScriptRoot\Types\Net45\HtmlAgilityPack.dll"
        } else {
            Add-Type -Path "$PSScriptRoot\Types\netstandard2.0\HtmlAgilityPack.dll"
        } 
    }
} catch {
    Write-Error -Message "Failed to load HtmlAgilityPack: $_"
    throw
}

$Classes = @(Get-ChildItem -Path $PSScriptRoot\Classes\*.ps1)
$Private = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1)
$Public = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1)

foreach ($ClassFile in $Classes) {
    try {
        . $ClassFile.FullName
    } catch {
        Write-Error -Message "Failed to import class $($ClassFile.FullName): $_"
        throw
    }
}

foreach ($Module in ($Private + $Public)) {
    try {
        . $Module.FullName
    } catch {
        Write-Error -Message "Failed to import function $($Module.FullName): $_"
        throw
    }
}

Export-ModuleMember -Function $Public.BaseName
