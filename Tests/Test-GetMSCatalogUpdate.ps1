Clear-Host

# Test-GetMSCatalogUpdate.ps1
# This script tests the key parameters of the Get-MSCatalogUpdate function

#region Module Import
# Import the module
$ModulePath = Split-Path $PSScriptRoot -Parent
Import-Module "$ModulePath\MSCatalogLTS.psd1" -Force
#endregion

#region Helper Functions
# Helper function to run tests and display results
function Test-Scenario {
    param (
        [string]$Name,
        [hashtable]$Parameters,
        [scriptblock]$Validation = { param($result) $result.Count -gt 0 }
    )
    
    Write-Host "`n======================================================" 
    Write-Host "Testing Scenario: $Name" -ForegroundColor Cyan
    Write-Host "======================================================" 
    
    # Build search query for display
    $searchText = if ($Parameters.ContainsKey('Search')) { 
        $Parameters.Search 
    } elseif ($Parameters.ContainsKey('OperatingSystem')) {
        $os = $Parameters.OperatingSystem
        $version = if ($Parameters.ContainsKey('Version')) { " Version $($Parameters.Version)" } else { "" }
        $updateType = if ($Parameters.ContainsKey('UpdateType')) { 
            if ($Parameters.UpdateType -is [array]) {
                " $($Parameters.UpdateType -join ' or ')"
            } else {
                " $($Parameters.UpdateType)"
            }
        } else { "" }
        "$os$version$updateType"
    } else {
        "custom search"
    }
    
    try {
        # Display parameter details
        Write-Host "Parameters:" -ForegroundColor Yellow
        $Parameters.GetEnumerator() | ForEach-Object {
            Write-Host "  $($_.Key): $($_.Value)" -ForegroundColor Gray
        }
        
        Write-Host "`nExecuting search for: $searchText..." -ForegroundColor Yellow
        
        # Execute the command and measure execution time
        $startTime = Get-Date
        $result = Get-MSCatalogUpdate @Parameters -ErrorAction Stop
        $executionTime = ((Get-Date) - $startTime).TotalSeconds
        
        if ($null -eq $result -or $result.Count -eq 0) {
            Write-Host "No results found for this query." -ForegroundColor Yellow
            $returnObj = New-Object PSObject -Property @{
                Success = $false
                Result = $null
                Message = "No results found"
                ExecutionTime = $executionTime
                SearchText = $searchText
            }
            return $returnObj
        }
        
        Write-Host "Found $($result.Count) updates in $([math]::Round($executionTime, 2)) seconds" -ForegroundColor Green
        
        # Display first few results
        if ($result.Count -gt 0) {
            Write-Host "`nSample Results:" -ForegroundColor Yellow
            $sampleSize = [Math]::Min(3, $result.Count)
            $result | Select-Object -First $sampleSize | ForEach-Object {
                Write-Host "- $($_.Title)" -ForegroundColor Gray
                Write-Host "  Classification: $($_.Classification)" -ForegroundColor Gray
            }
            
            if ($result.Count -gt $sampleSize) {
                Write-Host "- ... and $($result.Count - $sampleSize) more" -ForegroundColor Gray
            }
        }
        
        # Execute validation
        $validationResult = & $Validation $result
        if ($validationResult -eq $true) {
            Write-Host "`n[PASS] Test PASSED" -ForegroundColor Green
            $returnObj = New-Object PSObject -Property @{
                Success = $true
                Result = $result
                Message = "Test passed"
                ExecutionTime = $executionTime
                SearchText = $searchText
            }
            return $returnObj
        } 
        else {
            Write-Host "`n[FAIL] Test FAILED: Validation criteria not met" -ForegroundColor Red
            $returnObj = New-Object PSObject -Property @{
                Success = $false
                Result = $result
                Message = "Validation failed"
                ExecutionTime = $executionTime
                SearchText = $searchText
            }
            return $returnObj
        }
    } 
    catch {
        Write-Host "`n[ERROR] ERROR: $($_.Exception.Message)" -ForegroundColor Red
        $returnObj = New-Object PSObject -Property @{
            Success = $false
            Result = $null
            Message = $_.Exception.Message
            ExecutionTime = 0
            SearchText = $searchText
        }
        return $returnObj
    }
}

# Results tracking
$successCount = 0
$failureCount = 0
$testResults = @()

# Run test and record result
function Invoke-Test {
    param (
        [string]$Name,
        [hashtable]$Parameters,
        [scriptblock]$Validation = { param($result) $result.Count -gt 0 }
    )
    
    $result = Test-Scenario -Name $Name -Parameters $Parameters -Validation $Validation
    $testResult = New-Object PSObject -Property @{
        Name = $Name
        Parameters = ($Parameters.Keys -join ", ")
        Success = $result.Success
        ResultCount = if ($result.Result) { $result.Result.Count } else { 0 }
        ExecutionTime = [math]::Round($result.ExecutionTime, 2)
        Message = $result.Message
        SearchText = $result.SearchText
    }
    
    $script:testResults += $testResult
    
    if ($result.Success) {
        $script:successCount++
    } else {
        $script:failureCount++
    }
    
    return $result.Result
}
#endregion

#region Test Initialization
# Record start time for overall execution time
$scriptStartTime = Get-Date
#endregion

#region Test Group 1: Basic Search Parameter
# ======================================================
# TEST GROUP 1: Basic Search Parameter
# ======================================================
Write-Host "`n### TEST GROUP: BASIC SEARCH PARAMETER ###" -ForegroundColor Blue -BackgroundColor White

# Test 1.1: Basic KB number search
Invoke-Test -Name "KB Number Search" -Parameters @{
    Search = "KB5035853"
}

# Test 1.2: Basic Windows 11 search
Invoke-Test -Name "Windows 11 Search" -Parameters @{
    Search = "Cumulative Update for Windows 11 Version 24H2 for x64"
}

# Test 1.3: Basic Windows 10 search
Invoke-Test -Name "Windows 10 Search" -Parameters @{
    Search = "Cumulative Update for Windows 10 Version 22H2 for x64"
}

# Test 1.4: Cumulative updates search
Invoke-Test -Name "Cumulative Updates Search" -Parameters @{
    Search = "Cumulative Update"
}
#endregion

#region Test Group 2: OperatingSystem Parameter
# ======================================================
# TEST GROUP 2: OperatingSystem Parameter
# ======================================================
Write-Host "`n### TEST GROUP: OPERATING SYSTEM PARAMETER ###" -ForegroundColor Blue -BackgroundColor White

# Test 2.1: Windows 10 OS parameter with UpdateType
Invoke-Test -Name "Windows 10 OS Parameter" -Parameters @{
    OperatingSystem = "Windows 10"
    UpdateType = "Cumulative Updates"
} -Validation {
    param($result)
    ($result | Where-Object { $_.Title -match "Windows 10" }).Count -gt 0
}

# Test 2.2: Windows 11 OS parameter with UpdateType
Invoke-Test -Name "Windows 11 OS Parameter" -Parameters @{
    OperatingSystem = "Windows 11"
    UpdateType = "Cumulative Updates"
} -Validation {
    param($result)
    ($result | Where-Object { $_.Title -match "Windows 11" }).Count -gt 0
}

# Test 2.3: Windows Server OS parameter with UpdateType
Invoke-Test -Name "Windows Server OS Parameter" -Parameters @{
    OperatingSystem = "Windows Server"
    UpdateType = "Cumulative Updates"
} -Validation {
    param($result)
    ($result | Where-Object { $_.Title -match "Microsoft server operating system" }).Count -gt 0
}
#endregion

#region Test Group 3: OS + Version Parameter
# ======================================================
# TEST GROUP 3: OS + Version Parameter
# ======================================================
Write-Host "`n### TEST GROUP: OS + VERSION PARAMETER ###" -ForegroundColor Blue -BackgroundColor White

# Test 3.1: Windows 11 Version 22H2 with UpdateType
Invoke-Test -Name "Windows 11 + Version 22H2" -Parameters @{
    OperatingSystem = "Windows 11"
    Version = "22H2"
    UpdateType = "Cumulative Updates"
} -Validation {
    param($result)
    ($result | Where-Object { $_.Title -match "Windows 11.*(22H2|Version 22H2)" }).Count -gt 0
}

# Test 3.2: Windows 10 Version 21H2 with UpdateType
Invoke-Test -Name "Windows 10 + Version 21H2" -Parameters @{
    OperatingSystem = "Windows 10"
    Version = "21H2"
    UpdateType = "Cumulative Updates"
} -Validation {
    param($result)
    ($result | Where-Object { $_.Title -match "Windows 10.*(21H2|Version 21H2)" }).Count -gt 0
}
#endregion

#region Test Group 4: OS + UpdateType Parameter
# ======================================================
# TEST GROUP 4: OS + UpdateType Parameter
# ======================================================
Write-Host "`n### TEST GROUP: OS + UPDATE TYPE PARAMETER ###" -ForegroundColor Blue -BackgroundColor White

# Test 4.1: Windows 11 + Cumulative Updates
Invoke-Test -Name "Windows 11 + Cumulative Updates" -Parameters @{
    OperatingSystem = "Windows 11"
    UpdateType = "Cumulative Updates"
} -Validation {
    param($result)
    ($result | Where-Object { 
        $_.Title -match "Windows 11" -and $_.Title -match "Cumulative Update"
    }).Count -gt 0
}

# Test 4.2: Windows 11 + Security Updates 
# Add Cumulative Updates also to ensure results
Invoke-Test -Name "Windows 11 + Security Updates" -Parameters @{
    OperatingSystem = "Windows 11"
    UpdateType = @("Security Updates", "Cumulative Updates")
} -Validation {
    param($result)
    ($result | Where-Object { 
        $_.Title -match "Windows 11" -and 
        ($_.Title -match "Security Update|Cumulative Update" -or 
         $_.Classification -match "Security Updates|Cumulative Updates")
    }).Count -gt 0
}

# Test 4.3: Multiple update types
Invoke-Test -Name "Windows 11 + Multiple Update Types" -Parameters @{
    OperatingSystem = "Windows 11"
    UpdateType = @("Security Updates", "Cumulative Updates")
} -Validation {
    param($result)
    ($result | Where-Object { 
        $_.Title -match "Windows 11" -and
        ($_.Title -match "Cumulative Update|Security Update" -or 
         $_.Classification -match "Security Updates|Cumulative Updates")
    }).Count -gt 0
}
#endregion

#region Test Group 5: OS + Architecture Parameter
# ======================================================
# TEST GROUP 5: OS + Architecture Parameter 
# ======================================================
Write-Host "`n### TEST GROUP: OS + ARCHITECTURE PARAMETER ###" -ForegroundColor Blue -BackgroundColor White

# Test 5.1: Windows 11 + x64 with UpdateType
Invoke-Test -Name "Windows 11 + x64 Architecture" -Parameters @{
    OperatingSystem = "Windows 11"
    Architecture = "x64"
    UpdateType = "Cumulative Updates"
} -Validation {
    param($result)
    ($result | Where-Object { 
        $_.Title -match "Windows 11" -and
        $_.Title -match "x64-based|64-bit" -and -not ($_.Title -match "arm64|x86-based|32-bit")
    }).Count -gt 0
}

# Test 5.2: Windows 11 + ARM64 with UpdateType
Invoke-Test -Name "Windows 11 + ARM64 Architecture" -Parameters @{
    OperatingSystem = "Windows 11"
    Architecture = "arm64"
    UpdateType = "Cumulative Updates"
} -Validation {
    param($result)
    ($result | Where-Object { 
        $_.Title -match "Windows 11" -and 
        $_.Title -match "arm64|ARM-based"
    }).Count -gt 0
}
#endregion

#region Test Group 6: OS + Version + UpdateType
# ======================================================
# TEST GROUP 6: OS + Version + UpdateType
# ======================================================
Write-Host "`n### TEST GROUP: OS + VERSION + UPDATE TYPE ###" -ForegroundColor Blue -BackgroundColor White

# Test 6.1: Windows 11 + 22H2 + Cumulative Updates
Invoke-Test -Name "Windows 11 + 22H2 + Cumulative Updates" -Parameters @{
    OperatingSystem = "Windows 11"
    Version = "22H2"
    UpdateType = "Cumulative Updates"
} -Validation {
    param($result)
    ($result | Where-Object { 
        $_.Title -match "Windows 11.*22H2.*Cumulative Update" -or
        $_.Title -match "Cumulative Update.*Windows 11.*22H2"
    }).Count -gt 0
}
#endregion

#region Test Group 7: All Parameters Combined
# ======================================================
# TEST GROUP 7: All Parameters Combined
# ======================================================
Write-Host "`n### TEST GROUP: ALL PARAMETERS COMBINED ###" -ForegroundColor Blue -BackgroundColor White

# Test 7.1: OS + Version + UpdateType + Architecture
Invoke-Test -Name "All Parameters Combined" -Parameters @{
    OperatingSystem = "Windows 11"
    Version = "22H2"
    UpdateType = "Cumulative Updates"
    Architecture = "x64"
} -Validation {
    param($result)
    ($result | Where-Object { 
        ($_.Title -match "Windows 11.*22H2.*Cumulative Update.*x64-based" -or
         $_.Title -match "Cumulative Update.*Windows 11.*22H2.*x64-based") -and
        -not ($_.Title -match "arm64|x86-based|32-bit")
    }).Count -gt 0
}
#endregion

#region Test Summary
# Calculate script execution time
$scriptEndTime = Get-Date
$totalExecutionTime = ($scriptEndTime - $scriptStartTime).TotalSeconds

# Display test summary
$totalTests = $successCount + $failureCount
$successRate = if ($totalTests -gt 0) { [math]::Round(($successCount / $totalTests) * 100, 2) } else { 0 }

Write-Host "`n======================================================" 
Write-Host "TEST SUMMARY" -ForegroundColor Green
Write-Host "======================================================" 
Write-Host "Total Tests: $totalTests" -ForegroundColor Cyan
Write-Host "Successful: $successCount" -ForegroundColor Green
Write-Host "Failed: $failureCount" -ForegroundColor Red
Write-Host "Success Rate: $successRate%" -ForegroundColor Cyan
Write-Host "Total Execution Time: $([math]::Round($totalExecutionTime, 2)) seconds" -ForegroundColor Cyan

Write-Host "`nDetailed Test Results:" -ForegroundColor Cyan
$testResults | Format-Table Name, Parameters, Success, ResultCount, ExecutionTime, Message -AutoSize

Write-Host "`n======================================================" 
Write-Host "TEST COMPLETION" -ForegroundColor Green
Write-Host "======================================================"
#endregion
