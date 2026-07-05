#Requires -Version 5.1
<#
.SYNOPSIS
    Retest harness for the checkup.ps1 event-log fixes.
.DESCRIPTION
    Runs each OLD (buggy, provider-unpinned) Get-WinEvent query next to the
    CORRECTED (provider-pinned) query and prints counts side by side, plus the
    provider breakdown of what the buggy query was actually counting.
    Read-only: makes no changes to the system.
.EXAMPLE
    .\retest.ps1
#>

$ErrorActionPreference = 'Continue'

function Get-EventCount {
    param([hashtable]$Filter)
    try {
        $events = @(Get-WinEvent -FilterHashtable $Filter -ErrorAction Stop)
        return ,@($events)   # wrap so an empty array survives the pipeline
    } catch {
        if ($_.Exception.Message -match 'No events were found') { return ,@() }
        return $null   # real failure (access denied, bad log name, ...)
    }
}

function Show-Comparison {
    param(
        [string]$Name,
        [hashtable]$OldFilter,
        [hashtable]$NewFilter,
        [scriptblock]$NewPostFilter = $null
    )

    Write-Host ""
    Write-Host ("=== {0} {1}" -f $Name, ("=" * [Math]::Max(3, 55 - $Name.Length))) -ForegroundColor Cyan

    $old = Get-EventCount $OldFilter
    $new = Get-EventCount $NewFilter
    if ($null -ne $new -and $NewPostFilter) { $new = @($new | Where-Object $NewPostFilter) }

    $oldStr = if ($null -eq $old) { "QUERY FAILED" } else { "$($old.Count)" }
    $newStr = if ($null -eq $new) { "QUERY FAILED" } else { "$($new.Count)" }

    Write-Host ("  {0,-42} {1}" -f "OLD (buggy, unpinned):", $oldStr) -ForegroundColor Yellow
    Write-Host ("  {0,-42} {1}" -f "NEW (provider-pinned):", $newStr) -ForegroundColor Green

    if ($null -ne $old -and $old.Count -gt 0) {
        Write-Host "  What the OLD query was actually counting:" -ForegroundColor DarkGray
        $old | Group-Object ProviderName | Sort-Object Count -Descending | ForEach-Object {
            $marker = if ($_.Name -eq $NewFilter.ProviderName) { "(real)" } else { "(FALSE POSITIVE)" }
            Write-Host ("    {0,4} x {1} {2}" -f $_.Count, $_.Name, $marker) -ForegroundColor DarkGray
        }
    }

    if ($null -ne $old -and $null -ne $new -and $old.Count -ne $new.Count) {
        Write-Host ("  >> Buggy query miscounted by {0} event(s) on this machine." -f [math]::Abs($old.Count - $new.Count)) -ForegroundColor Magenta
    }
}

Write-Host ""
Write-Host ("checkup.ps1 event-query retest - {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -ForegroundColor White
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
Write-Host ("CPU: {0}" -f $cpu.Name.Trim()) -ForegroundColor DarkGray

$last7  = (Get-Date).AddDays(-7)
$last30 = (Get-Date).AddDays(-30)

# 1. Thermal throttling (the originally reported bug)
Show-Comparison -Name "Thermal throttling (System, Id 37/55)" `
    -OldFilter @{ LogName = 'System'; Id = 37; StartTime = $last7 } `
    -NewFilter @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Processor-Power'; Id = @(37, 55); StartTime = $last7 } `
    -NewPostFilter { $_.Level -ge 1 -and $_.Level -le 3 }

# Show why the Level filter on Id 55 matters
$kpp55 = Get-EventCount @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Processor-Power'; Id = 55; StartTime = $last7 }
if ($null -ne $kpp55 -and $kpp55.Count -gt 0) {
    $info55 = @($kpp55 | Where-Object { $_.Level -gt 3 -or $_.Level -eq 0 }).Count
    Write-Host ("  Note: {0} KPP Id=55 event(s) exist; {1} are informational per-core" -f $kpp55.Count, $info55) -ForegroundColor DarkGray
    Write-Host "  boot-time capability dumps excluded by the Level filter." -ForegroundColor DarkGray
}

# 2. App crashes
Show-Comparison -Name "App crashes (Application, Id 1000)" `
    -OldFilter @{ LogName = 'Application'; Id = 1000; StartTime = $last7 } `
    -NewFilter @{ LogName = 'Application'; ProviderName = 'Application Error'; Id = 1000; StartTime = $last7 }

# 3. Blue screens
Show-Comparison -Name "Blue screens (System, Id 1001)" `
    -OldFilter @{ LogName = 'System'; Id = 1001; StartTime = $last30 } `
    -NewFilter @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-WER-SystemErrorReporting'; Id = 1001; StartTime = $last30 }

# 4. Live clock ratio (runtime throttle confirmation signal)
Write-Host ""
Write-Host "=== Live clock ratio (runtime signal) ==================" -ForegroundColor Cyan
try {
    $perf = [math]::Round((Get-Counter '\Processor Information(_Total)\% Processor Performance' -ErrorAction Stop).CounterSamples[0].CookedValue, 0)
    Write-Host ("  % Processor Performance (counter):        {0}%  (boost can exceed 100)" -f $perf) -ForegroundColor Green
} catch {
    Write-Host "  Perf counter unavailable (localized name?) - WMI fallback only" -ForegroundColor Yellow
}
if ($cpu.MaxClockSpeed -gt 0) {
    $wmiPct = [math]::Round(($cpu.CurrentClockSpeed / $cpu.MaxClockSpeed) * 100, 0)
    Write-Host ("  Win32_Processor Current/Max:               {0}/{1} MHz ({2}%) - often pinned to base clock, fallback only" -f $cpu.CurrentClockSpeed, $cpu.MaxClockSpeed, $wmiPct) -ForegroundColor DarkGray
}

# 5. MSAcpi_ThermalZoneTemperature availability (should be info, never a failure)
Write-Host ""
Write-Host "=== ACPI thermal zone (should degrade to 'info') =======" -ForegroundColor Cyan
try {
    $tz = @(Get-CimInstance MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" -ErrorAction Stop)
    $valid = @($tz | ForEach-Object { [math]::Round(($_.CurrentTemperature - 2732) / 10, 1) } | Where-Object { $_ -gt 0 -and $_ -lt 120 })
    if ($valid.Count -gt 0) {
        Write-Host ("  Available: {0} C ({1} valid zone(s))" -f ($valid -join ' C, '), $valid.Count) -ForegroundColor Green
    } else {
        Write-Host "  Class present but no plausible readings - checkup reports 'info', correct." -ForegroundColor Green
    }
} catch {
    Write-Host "  Unavailable on this machine (typical for desktop boards) - checkup" -ForegroundColor Green
    Write-Host "  reports 'info/not available', NOT a failure. Correct behavior." -ForegroundColor Green
}

Write-Host ""
Write-Host "Done. If OLD and NEW counts differ above, that difference is exactly" -ForegroundColor White
Write-Host "the false-positive noise the provider pinning removes." -ForegroundColor White
Write-Host ""
