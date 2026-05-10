function Get-UpdateSupportUrl {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Guid
    )

    $Uri = "https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=$Guid"

    try {
        $Resp = Invoke-WebRequest -Uri $Uri -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Verbose "Failed to fetch detail page for $Guid : $_"
        return ""
    }

    $Doc = [HtmlAgilityPack.HtmlDocument]::new()
    $Doc.LoadHtml($Resp.RawContent.ToString())

    $Node = $Doc.GetElementbyId("suportUrlDiv")
    if ($null -eq $Node) { return "" }

    $InnerDivs = $Node.SelectNodes("div")
    if ($null -eq $InnerDivs -or $InnerDivs.Count -eq 0) { return "" }

    return $InnerDivs[0].InnerText.Trim()
}
