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
        [switch] $GetFramework,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("all", "x64", "x86", "arm64")]
        [string] $Architecture = "all",
        
        [Parameter(Mandatory = $false)]
        [switch] $IncludePreview,
        
        [Parameter(Mandatory = $false)]
        [switch] $IncludeDynamic
    )

   try {
       $ProgPref = $ProgressPreference
       $ProgressPreference = "SilentlyContinue"

        $Rows = @() 
        $PageCount = 0

        if ($Strict) {
            $EncodedSearch = [uri]::EscapeDataString('"' + $Search + '"') 
        } elseif ($GetFramework){
            $EncodedSearch = [uri]::EscapeDataString("*$Search*")
        } else {
            $EncodedSearch = [uri]::EscapeDataString("$Search")
        }
    
        $Uri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$EncodedSearch"
        $Res = Invoke-CatalogRequest -Uri $Uri
        $Rows = $Res.Rows

        if ($AllPages) {
            while ($Res.NextPage -and $PageCount -lt 39) { # 40 pages is the limit
                $PageCount++
                $All = "$Uri&p=$PageCount"
                $Res = Invoke-CatalogRequest -Uri $All
                $Rows += $Res.Rows
                }
            } 

        if (-not $IncludeDynamic) {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -notlike "*Dynamic*"})  
            }

        if (-not $IncludePreview) {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -notlike "*Preview*"})  
            }

        if ($ExcludeFramework) {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -notlike "*Framework*"})  
            }

        if ($Architecture -ne "all") {
            switch ($Architecture) {
                "x64" { 
                    $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*x64*" -or $_.SelectNodes("td")[1].InnerText.Trim() -like "*64-Bit*"})
                }
                "x86" {
                    $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*x86*" -or $_.SelectNodes("td")[1].InnerText.Trim() -like "*32-Bit*"})
                }
                "arm64" {
                    $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*arm64*" -or $_.SelectNodes("td")[1].InnerText.Trim() -like "*ARM*"})
                }
            }
        }

        if ($Search -match "Windows 10") {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*Windows 10*"})  
            }

        if ($Search -match "Windows 11") {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*Windows 11*"})  
            }

        if ($Search -match "Windows Server") {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*Windows Server*"})  
            }

        if ($GetFramework) {
            $Rows = $Rows.Where({$_.SelectNodes("td")[1].InnerText.Trim() -like "*Framework*" -or $_.SelectNodes("td")[1].InnerText.Trim() -like "*.NET Framework*"}) 
            }

        Write-Host "`nFiltered search completed. Total rows: $($Rows.Count)"

        if ($Rows.Count -ge 1000) {
            Write-Host "`nMax Result limit of 1000 hit, please refine your search criteria"
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