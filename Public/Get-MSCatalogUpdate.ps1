function Get-MSCatalogUpdate {  
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Search,

        [Parameter(Mandatory = $false)]
        [switch] $IncludeFileNames,

        [Parameter(Mandatory = $false)]
        [switch] $AllPages,

        [Parameter(Mandatory = $false)]
        [switch] $ExcludeFramework,

        [Parameter(Mandatory = $false)]
        [switch] $Strict,

        [Parameter(Mandatory = $false)]
        [switch] $GetFramework
    )
    
# Default settings for the search
    
$Bit64 = $true  # Include only x64 updates
$ExcludePreview = $true # Exclude Preview updates
$ExcludeDynamic = $true # Exclude Dynamic updates


   try {
       $ProgPref = $ProgressPreference
       $ProgressPreference = "SilentlyContinue"

        $Rows = @() 
        $PageCount = 0

        $Uri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$([uri]::EscapeDataString($Search))"
        $Res = Invoke-CatalogRequest -Uri $Uri
        $Rows = $Res.Rows

        if ($AllPages) {
            while ($Res.NextPage -and $PageCount -lt 3) {
                $PageCount++
                $All = "$Uri&p=$PageCount"
                $Res = Invoke-CatalogRequest -Uri $All
                $Rows += $Res.Rows
                }
            } 

        if ($Strict) { 
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*$Search*"})
            }

        if ($ExcludeDynamic) {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -notlike "*Dynamic*"})  
            }

        if ($ExcludePreview) {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -notlike "*Preview*"})  
            }

        if ($ExcludeFramework) {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -notlike "*Framework*"})  
            }

        if ($Bit64) {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*x64*" -or $_.SelectNodes("td")[1].InnerText.Trim() -like "*64-Bit*"})  
            }

        if ($Search -match "Windows 10") {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*Windows 10*"})  
            }

        if ($Search -match "Windows 11") {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*Windows 11*"})  
            }

        if ($GetFramework) {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*Framework*"})  
            }

       if ($Rows.Count -gt 0) {
           foreach ($Row in $Rows) {
               if ($Row.Id -ne "headerRow") {
                   [MSCatalogUpdate]::new($Row, $IncludeFileNames)
               }
           }
       } else {
           Write-Warning "No updates found matching the search term."
       }
   } catch {
       if ($_.Exception.Message -like "No updates found matching*") {
           Write-Warning "No updates found matching the search term."
       } else {
           Write-Warning "We did not find any results for $Search"
       }
       $ProgressPreference = $ProgPref
   }
}