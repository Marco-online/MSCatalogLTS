# MSCatalogLTS

MSCatalogLTS is a Long-term support module for searching and downloading updates from https://www.catalog.update.microsoft.com.
It is cross-platform and runs on both Desktop and Core versions of PowerShell.

[![psgallery](https://img.shields.io/powershellgallery/v/mscataloglts?style=flat-square&logo=powershell)](https://www.powershellgallery.com/packages/MSCatalogLTS)

## Importing the Module

Import the module MSCatalogLTS direct from PSgallery:

``` powershell
Install-Module -Name MSCatalogLTS
```

If you already have the module installed, make sure it is the latest version:
``` powershell
Update-Module -Name MSCatalogLTS
```
## Get-MSCatalogUpdate

Use the Get-MSCatalogUpdate command to retrieve updates from the Microsoft Update Catalog. 
By default, this command returns the first 25 items from the search, sorted by the LastUpdated field in descending order.  
AllPages parameter will return all (max 1000) available results.

Retrieve cumulative updates for Windows 11 Version 24H2

```powershell
Get-MSCatalogUpdate -AllPages -Search "Cumulative Update for Windows 11 Version 24H2 for x64" -Strict
```
```powershell
Title                                                                                               Products   Classification   LastUpdated Size    
-----                                                                                               --------   --------------   ----------- ----    
2024-07 Cumulative Update for Windows 11 Version 24H2 for x64-based Systems (KB5040435)             Windows 11 Security Updates 2024/07/09  302.0 MB
2024-06 Cumulative Update for Windows 11 Version 24H2 for x64-based Systems (KB5039239)             Windows 11 Security Updates 2024/06/15  248.8 MB
```

Retrieve only .NET Framework updates for Windows 11 Version 24H2:

```powershell
Get-MSCatalogUpdate -Search "Cumulative Update for Windows 11 Version 24H2" -GetFramework
```
```powershell
Title                                                                                                       Products   Classification   LastUpdated Size    
-----                                                                                                       --------   --------------   ----------- ---- 
2024-07 Cumulative Update for .NET Framework 3.5 and 4.8.1 for Windows 11, version 24H2 for x64 (KB5039894) Windows 11 Security Updates 2024/07/09  70.9 MB
```
## Filtering and Multi-Link Support

```powershell
Get-MSCatalogUpdate
-AllPages               Get results of maximum 1000 hits.
-Architecture           Allows filtering updates by architecture (all, x64, x86, arm64).
-IncludePreview         Includes or excludes preview updates.
-IncludeDynamic         Includes or excludes dynamic updates.
-Strict                 Get results which only contain the exact search term.
-ExcludeFramework       Hide .NET Framework updates from Cumulative updates results.
-GetFramework           Only show .NET Framework updates.

Save-MSCatalogUpdate
-DownloadAll            Multiple download links associated with a single update.

```
Examples:
```powershell
Get-MSCatalogUpdate -AllPages -Search "Cumulative Update for Windows 11 Version 24H2 for x64" -Strict
Get-MSCatalogUpdate -AllPages -Search "Cumulative Update for Windows 10" -Architecture x64 -ExcludeFramework
Get-MSCatalogUpdate -Search "Cumulative Update for Windows 11 Version 24H2" -ExcludeFramework
```
## Save-MSCatalogUpdate

Use the Save-MSCatalogUpdate command to download update files from the Microsoft Update Catalog.
Download updates specified by an object returned from Get-MSCatalogUpdate:

```powershell
$update = Get-MSCatalogUpdate -Search "Cumulative Update for Windows 11 Version 24H2" -ExcludeFramework
Save-MSCatalogUpdate -update $update -Destination ".\" -DownloadAll (-DownloadAll is an example to download all files if any)
```
## Save-MSCatalogOutput

The Save-MSCatalogOutput function allows you to save the output from the Get-MSCatalogUpdate command to an Excel file. This is particularly useful for maintaining a record of updates retrieved from the Microsoft Update Catalog.

Before using the Save-MSCatalogOutput function, ensure that you have the ImportExcel module installed. You can install it using the following command:
```powershell
Install-Module -Name ImportExcel
```
By default, the worksheet name in the Excel file will be named "Updates". However, you can customize the worksheet name by specifying the -WorksheetName parameter.

```powershell
Save-MSCatalogOutput -Update $update -WorksheetName "08_2024_Updates" -Destination "C:\Temp\2024_Updates.xlsx"
```

## HtmlAgilityPack

MSCatalogLTS uses the HtmlAgilityPack library for HTML parsing to ensure cross-platform compatibility. This avoids reliance on the Windows-only ParsedHtml property of the Invoke-WebRequest CmdLet.

## Credits

MSCatalogLTS is heavily inspired by MSCatalog. Some parts of MSCatalogLTS's code are plain copies. Big thanks to the MSCatalog team for creating!
Archive : https://github.com/ryan-jan/MSCatalog