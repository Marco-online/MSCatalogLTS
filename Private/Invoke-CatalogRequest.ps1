function Invoke-CatalogRequest {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [string] $Uri,

        [Parameter(Mandatory = $false)]
        [string] $Method = "Get",

        [Parameter(Mandatory = $false)]
        [string] $EventArgument,

        [Parameter(Mandatory = $false)]
        [string] $EventTarget,

        [Parameter(Mandatory = $false)]
        [string] $EventValidation,

        [Parameter(Mandatory = $false)]
        [string] $ViewState,

        [Parameter(Mandatory = $false)]
        [string] $ViewStateGenerator,

        [switch] $ShowDebug
    )

    try {
        Set-TempSecurityProtocol

        if ($Method -eq "Post") {
            $ReqBody = @{
                "__EVENTARGUMENT" = $EventArgument
                "__EVENTTARGET" = $EventTarget
                "__EVENTVALIDATION" = $EventValidation
                "__VIEWSTATE" = $ViewState
                "__VIEWSTATEGENERATOR" = $ViewStateGenerator
            }
        }
        $Params = @{
            Uri = $Uri
            Method = $Method
            Body = $ReqBody
            ContentType = "application/x-www-form-urlencoded"
            UseBasicParsing = $true
            ErrorAction = "Stop"
        }

        $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        if ($Stopwatch.Elapsed.TotalSeconds -ge 60) {
            Write-Warning "Timeout reached (60 seconds)"
            Set-TempSecurityProtocol -ResetToDefault
            return
        }

        $Results = Invoke-WebRequest @Params
        $Stopwatch.Stop()

        if ($ShowDebug) {
            Write-Host "DEBUG: Request $Uri" -ForegroundColor yellow 
            Write-Host "DEBUG: Request completed in $($Stopwatch.Elapsed.TotalSeconds) seconds" -ForegroundColor yellow            
        }

        $HtmlDoc = [HtmlAgilityPack.HtmlDocument]::new()
        $HtmlDoc.LoadHtml($Results.RawContent.ToString())
        $NoResults = $HtmlDoc.GetElementbyId("ctl00_catalogBody_noResultText")
        if ($null -eq $NoResults) {
            $ErrorText = $HtmlDoc.GetElementbyId("errorPageDisplayedError")
            if ($ErrorText) {
                throw "The catalog.microsoft.com site has encountered an error. Please try again later."
            } else {
                [MSCatalogResponse]::new($HtmlDoc)
            }
        }       
            } catch {
                Write-Warning "$_"
            } finally {
                Set-TempSecurityProtocol -ResetToDefault
            }
        }