<#
    .SYNOPSIS
        This command is used to retrieve updates from the https://www.catalog.update.microsoft.com website.

    .EXAMPLE
        $update = Get-MSCatalogUpdate -AllPages -Search "Cumulative Update for Windows 10 Version 21H2" -GetFramework
#>
function Get-MSCatalogUpdate {  
    [CmdLetBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Search,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Title", "Products", "Classification", "LastUpdated", "Size")]
        [string] $SortBy,

        [Parameter(Mandatory = $false)]
        [switch] $Descending,

        [Parameter(Mandatory = $false)]
        [switch] $Strict,

        [Parameter(Mandatory = $false)]
        [switch] $IncludeFileNames,

        [Parameter(Mandatory = $false)]
        [switch] $AllPages,

        [Parameter(Mandatory = $false)]
        [switch] $ExcludeFramework,

        [Parameter(Mandatory = $false)]
        [switch] $GetFramework,
              
        [Parameter(Mandatory = $false)]
        [switch] $Version10,
                
        [Parameter(Mandatory = $false)]
        [switch] $Version11                   
    )

# Default settings for the search
    
    $Class = $true # Include only Security Updates
    $Bit64 = $true  # Include only x64 updates
    $ExcludePreview = $true # Exclude Preview updates
    $ExcludeDynamic = $true # Exclude Dynamic updates

    try {
        $ProgPref = $ProgressPreference
        $ProgressPreference = "SilentlyContinue"

        $Uri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$([uri]::EscapeDataString($Search))"
        $Res = Invoke-CatalogRequest -Uri $Uri

        if ($PSBoundParameters.ContainsKey("SortBy")) {
            $SortParams = @{
                Uri = $Uri
                SortBy = $SortBy
                Descending = $Descending
                EventArgument = $Res.EventArgument
                EventValidation = $Res.EventValidation
                ViewState = $Res.ViewState
                ViewStateGenerator = $Res.ViewStateGenerator
            }
            $Res = Sort-CatalogResults @SortParams
        } else {
            # Default sort is by LastUpdated and in descending order.
            $SortParams = @{
                Uri = $Uri
                SortBy = "LastUpdated"
                Descending = $true
                EventArgument = $Res.EventArgument
                EventValidation = $Res.EventValidation
                ViewState = $Res.ViewState
                ViewStateGenerator = $Res.ViewStateGenerator
            }
            $Res = Sort-CatalogResults @SortParams
        }

        $Rows = $Res.Rows

        if ($Strict -and -not $AllPages) {
            $StrictRows = $Rows.Where({
                $_.SelectNodes("td")[1].InnerText.Trim() -like "*$Search*"
            })
            while (($StrictRows.Count -lt 25) -and ($Res.NextPage -eq "")) {
                $NextParams = @{
                    Uri = $Uri
                    EventArgument = $Res.EventArgument
                    EventTarget = 'ctl00$catalogBody$nextPageLinkText'
                    EventValidation = $Res.EventValidation
                    ViewState = $Res.ViewState
                    ViewStateGenerator = $Res.ViewStateGenerator
                    Method = "Post"
                }
                $Res = Invoke-CatalogRequest @NextParams
                $StrictRows += $Res.Rows.Where({
                    $_.SelectNodes("td")[1].InnerText.Trim() -like "*$Search*"
                })
            }
            $Rows = $StrictRows[0..24]
        } elseif ($AllPages) {
            while ($Res.NextPage -eq "") {
                $NextParams = @{
                    Uri = $Uri
                    EventArgument = $Res.EventArgument
                    EventTarget = 'ctl00$catalogBody$nextPageLinkText'
                    EventValidation = $Res.EventValidation
                    ViewState = $Res.ViewState
                    ViewStateGenerator = $Res.ViewStateGenerator
                    Method = "Post"
                }
                $Res = Invoke-CatalogRequest @NextParams
                $Rows += $Res.Rows
            }
            if ($Strict) {
                $Rows = $Rows.Where({
                    $_.SelectNodes("td")[1].InnerText.Trim() -like "*$Search*"
                })
            }
        }

        if ($ExcludePreview) {
            $Rows = $Rows | Where-Object {
                $nodes = $_.SelectNodes("td")
                if ($nodes -and $nodes.Count -gt 1) {
                    $nodes[1].InnerText.Trim() -notlike "*Preview*"
                } else {
                    $false
                }
            }
        }

        if ($ExcludeDynamic) {
            $Rows = $Rows | Where-Object {
                $nodes = $_.SelectNodes("td")
                if ($nodes -and $nodes.Count -gt 1) {
                    $nodes[1].InnerText.Trim() -notlike "*Dynamic*"
                } else {
                    $false
                }
            }
        }

        if ($Bit64) {
            $Rows = $Rows | Where-Object {
                $nodes = $_.SelectNodes("td")
                if ($nodes -and $nodes.Count -gt 1) {
                    $nodes[1].InnerText.Trim() -like "*x64*"
                } else {
                    $false
                }
            }
        }

        if ($Version10) {
            $Rows = $Rows | Where-Object {
                $nodes = $_.SelectNodes("td")
                if ($nodes -and $nodes.Count -gt 1) {
                    $nodes[1].InnerText.Trim() -like "*Windows 10*"
                } else {
                    $false
                }
            }
        }

        if ($Version11) {
            $Rows = $Rows | Where-Object {
                $nodes = $_.SelectNodes("td")
                if ($nodes -and $nodes.Count -gt 1) {
                    $nodes[1].InnerText.Trim() -like "*Windows 11*"
                } else {
                    $false
                }
            }
        }

        if ($GetFramework) {
            $Rows = $Rows | Where-Object {
                $nodes = $_.SelectNodes("td")
                if ($nodes -and $nodes.Count -gt 1) {
                    $nodes[1].InnerText.Trim() -like "*4.8*" -and $nodes[1].InnerText.Trim() -match "Framework"
                } else {
                    $false
                }
            }
        }
        
        if ($Class) {
            $Rows = $Rows | Where-Object {
                $nodes = $_.SelectNodes("td")
                if ($nodes -and $nodes.Count -gt 3) {
                    $nodes[3].InnerText.Trim() -like "*Security Updates*"
                } else {
                    $false
                }
            }
        }

        if ($ExcludeFramework) {
            $Rows = $Rows | Where-Object {
                $nodes = $_.SelectNodes("td")
                if ($nodes -and $nodes.Count -gt 1) {
                    $nodes[1].InnerText.Trim() -notmatch "Framework"
                } else {
                    $false
                }
            }
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
        $ProgressPreference = $ProgPref
    } catch {
        $ProgressPreference = $ProgPref
        if ($_.Exception.Message -like "We did not find*") {
            Write-Warning $_.Exception.Message
        } else {
            throw $_
        }
    }
}
