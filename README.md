# MSCatalogLTS 
PowerShell Module for Windows Updates

[![PSGallery Version](https://img.shields.io/powershellgallery/v/MSCatalogLTS.svg?style=flat&logo=powershell&label=PSGallery%20Version)](https://www.powershellgallery.com/packages/MSCatalogLTS) [![PSGallery Downloads](https://img.shields.io/powershellgallery/dt/MSCatalogLTS.svg?style=flat&logo=powershell&label=PSGallery%20Downloads)](https://www.powershellgallery.com/packages/MSCatalogLTS) [![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?style=flat&logo=powershell)](https://www.powershellgallery.com/packages/MSCatalogLTS) [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> **Long-term support PowerShell module** for searching, filtering, and downloading Windows updates from the Microsoft Update Catalog

MSCatalogLTS is a **cross-platform PowerShell module** that provides an easy-to-use interface for interacting with the [Microsoft Update Catalog](https://www.catalog.update.microsoft.com). Whether you need to download cumulative updates, security patches, or manage offline updates for Windows 10, Windows 11, or Windows Server, MSCatalogLTS makes it simple.

---

## ⚡Key Features

- 🔍 **Search Windows Updates** - Find cumulative updates, security patches, drivers, and more
- ⬇️ **Download Updates Offline** - Download .msu, .cab, and other update files directly
- 🖥️ **Cross-Platform Support** - Works on Windows PowerShell 5.1+ and PowerShell 7+ (Core)
- 🧠 **Smart Filtering** - Filter by date, size, architecture (x64, x86, arm64), product, and classification
- 📦 **Export to Multiple Formats** - Export results to CSV, JSON, XML, or Excel
- 🔧 **Pipeline Support** - Full PowerShell pipeline integration for automation
- 🛡️ **Long-Term Support** - Actively maintained with regular updates and bug fixes
- 📊 **Batch Operations** - Process multiple updates efficiently

---

## 📦 Quick Install

### Install from PowerShell Gallery

```powershell
Install-Module -Name MSCatalogLTS -Scope CurrentUser
```

### Update to Latest Version

```powershell
Update-Module -Name MSCatalogLTS
```

### Verify Installation

```powershell
Get-Module -ListAvailable MSCatalogLTS
```

---

## 💡 Quick Start Examples

### Example 1: Search for Windows 11 Updates

```powershell
Get-MSCatalogUpdate -Search "Windows 11 24H2 x64"
```

**Output:**
```
Search completed for: Cumulative Update for Windows 11 Version 24H2 for x64
Found 2 updates

Title                                                                                                Products   Classification   LastUpdated Size
-----                                                                                                --------   --------------   ----------- ----
2026-01 Cumulative Update for Windows 11 Version 24H2 for x64-based Systems (KB5078127) (26100.7628) Windows 11 Updates 2026/01/24  4252.6 MB
2026-01 Cumulative Update for Windows 11 Version 24H2 for x64-based Systems (KB5077744) (26100.7627) Windows 11 Updates 2026/01/17  4252.3 MB
```

### Example 2: Download an Update

```powershell
# Search and download in one pipeline
Get-MSCatalogUpdate -Search "KB5065426" | Save-MSCatalogUpdate -Destination "C:\Updates"
```

### Example 3: Filter by Date Range

```powershell
Get-MSCatalogUpdate -Search "Windows 11 24H2 x64" -FromDate "2025-06-01" -ToDate "2025-07-01"
```
---



### Example 4: Save an Update

```powershell
# Get update and download
$update = Get-MSCatalogUpdate -Search "Windows 11 24H2 x64" | Select-Object -First 1
Save-MSCatalogUpdate -Update $update -Destination "C:\Updates" -DownloadAll

# One-liner
Get-MSCatalogUpdate -Search "Windows 11 24H2 x64" | Select-Object -First 1 | Save-MSCatalogUpdate -Destination "C:\Updates" -DownloadAll
```

### Example 5: Export to JSON

```powershell
# Export to JSON (overwrite)
Get-MSCatalogUpdate -Search "Windows Server 2022 x64" -LastDays 30 -ExportJson "C:\Updates\Updates.json"

# Export to JSON (append)
Get-MSCatalogUpdate -Search "Windows Server 2022 x64" -ExportJson "C:\Updates\Updates.json" -Append

When -Append is used, existing entries are preserved and duplicate updates are automatically skipped
Output: Appended 2 updates to JSON file: updates.json (skipped 23 duplicates)

```

---

## 📚 Available Functions

| Function | Description |
|----------|-------------|
| `Get-MSCatalogUpdate` | Search and filter updates from Microsoft Update Catalog |
| `Save-MSCatalogUpdate` | Download update files to a specified location |
| `Save-MSCatalogOutput` | Export update information to Excel format |

---

## 🔍 Common Use Cases

### Windows 10/11 Patch Management
```powershell
# Get latest Windows 11 security updates
Get-MSCatalogUpdate -Search "Windows 11 24h2 x64" -UpdateType "Security Updates" -LastDays 30
```

### WSUS/SCCM Offline Updates
```powershell
# Download updates for offline deployment
Get-MSCatalogUpdate -Search "KB5043178" -Architecture x64 | Save-MSCatalogUpdate -Destination "\\server\updates" -DownloadAll
```

### Server Patching Automation
```powershell
# Get Windows Server 2022 updates from last 90 days
Get-MSCatalogUpdate -Search "Windows Server 2022" -LastDays 90
```

### Filter by Size for Bandwidth Management
```powershell
# Find updates smaller than 500MB
Get-MSCatalogUpdate -Search "Windows 10 1809" -MaxSize 500 -SizeUnit MB
```

---

## 📖 Full Documentation

For detailed documentation, advanced examples, and parameter reference, visit:

👉 **[MSCatalogLTS Wiki](https://github.com/Marco-online/MSCatalogLTS/wiki/MSCatalogLTS)**

👉 **[GitHub Pages Documentation](https://marco-online.github.io/MSCatalogLTS)** *(coming soon)*

---

## 🛠️ Requirements

- **PowerShell**: 5.1 or higher (including PowerShell 7+)
- **Operating System**: Windows 10/11, Windows Server 2016+, or any OS with PowerShell Core
- **Internet Connection**: Required to query Microsoft Update Catalog
- **Optional**: ImportExcel module (for `Save-MSCatalogOutput` function)

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit:

- 🐛 Bug reports via [Issues](https://github.com/Marco-online/MSCatalogLTS/issues)
- 💡 Feature requests
- 🔧 Pull requests

---

## 📝 License

This project is licensed under the [MIT License](LICENSE).

---

## 🌟 Credits

MSCatalogLTS is inspired by and builds upon [MSCatalog](https://github.com/ryan-jan/MSCatalog) by ryan-jan. Thank you for the original groundwork!

---

## 🏷️ Keywords

`powershell` `windows-update` `microsoft-update-catalog` `patch-management` `wsus` `sccm` `windows-10` `windows-11` `windows-server` `security-updates` `cumulative-updates` `offline-updates` `powershell-module` `system-administration` `it-automation` `update-management` `kb-updates` `msu-files` `cab-files` `enterprise-deployment`

---

## 📧 Support

- 📖 [Documentation](https://github.com/Marco-online/MSCatalogLTS/wiki)
- 💬 [Issues](https://github.com/Marco-online/MSCatalogLTS/issues)
- 🌟 [PowerShell Gallery](https://www.powershellgallery.com/packages/MSCatalogLTS)

---

<p align="center">
  <strong>⭐ If you find MSCatalogLTS useful, please consider giving it a star!</strong>
</p>

<p align="center">
  Made with ❤️ for the PowerShell community
</p>