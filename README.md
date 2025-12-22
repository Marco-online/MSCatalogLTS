# MSCatalogLTS

MSCatalogLTS is a Long-term support module for searching and downloading updates from https://www.catalog.update.microsoft.com.
It is cross-platform and runs on both Desktop and Core versions of PowerShell.

[![psgallery](https://img.shields.io/powershellgallery/v/mscataloglts?style=flat-square&logo=powershell)](https://www.powershellgallery.com/packages/MSCatalogLTS)


# MSCatalogLTS
[![PSGallery Version](https://img.shields.io/powershellgallery/v/MSCatalogLTS.svg?style=flat&logo=powershell&label=PSGallery%20Version)](https://www.powershellgallery.com/packages/MSCatalogLTS) [![PSGallery Downloads](https://img.shields.io/powershellgallery/dt/MSCatalogLTS.svg?style=flat&logo=powershell&label=PSGallery%20Downloads)](https://www.powershellgallery.com/packages/MSCatalogLTS) [![PowerShell](https://img.shields.io/badge/PowerShell-5.1-blue?style=flat&logo=powershell)](https://www.powershellgallery.com/packages/MSCatalogLTS) [![PSGallery Platform](https://img.shields.io/badge/platform-Windows-blue?logo=windows)](https://www.powershellgallery.com/packages/MSCatalogLTS)

## Quick Install

``` powershell
Install-Module -Name MSCatalogLTS
```

Update to the latest version:

```powershell
Update-Module -Name MSCatalogLTS
```

```powershell
Get-MSCatalogUpdate -Search "Cumulative Update for Windows 11 Version 24H2 for x64" -Strict -LastDays 60 -Descending

Search completed for: Cumulative Update for Windows 11 Version 24H2 for x64
Found 2 updates

Title                                                                                                Products   Classification   LastUpdated Size
-----                                                                                                --------   --------------   ----------- ----
2025-09 Cumulative Update for Windows 11 Version 24H2 for x64-based Systems (KB5065426) (26100.6584) Windows 11 Security Updates 2025/09/09  3811.1 MB
2025-08 Cumulative Update for Windows 11 Version 24H2 for x64-based Systems (KB5063878) (26100.4946) Windows 11 Security Updates 2025/08/12  3054.9 MB
```



---

## 📚 Full Documentation

Looking for usage examples, parameter reference, or advanced filtering?

👉 Visit the [MSCatalogLTS Wiki](https://github.com/Marco-online/MSCatalogLTS/wiki/MSCatalogLTS)

---

## Credits

Inspired by [MSCatalog](https://github.com/ryan-jan/MSCatalog) — thanks to the original author!


