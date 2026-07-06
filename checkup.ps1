#Requires -Version 5.1
<#
.SYNOPSIS
    PC Checkup - Windows notebook health & performance diagnostic tool.
.DESCRIPTION
    Scans system hardware, power settings, memory pressure, startup bloat,
    security overhead, thermals, and common performance drains.
    Can export a plain-text report for AI analysis or pipe directly to
    a locally installed AI CLI.
.PARAMETER Report
    Output a clean plain-text report (no colors) to stdout. Ideal for piping
    or redirecting to a file.
.PARAMETER Clipboard
    Copy the plain-text report to clipboard for pasting into ChatGPT, Claude, etc.
.PARAMETER Analyze
    Pipe the report to a locally installed AI CLI for instant analysis.
    Auto-detects: claude, codex, ollama, aichat, sgpt, fabric, mods, etc.
.PARAMETER Ai
    Force a specific AI CLI for -Analyze (e.g. -Ai ollama).
.PARAMETER Model
    Model override for AI CLIs that support it (e.g. -Model llama3).
.PARAMETER Save
    Save the plain-text report to a file (e.g. -Save report.txt).
.EXAMPLE
    .\checkup.ps1
    .\checkup.ps1 -Clipboard
    .\checkup.ps1 -Analyze
    .\checkup.ps1 -Analyze -Ai ollama -Model llama3
    .\checkup.ps1 -Report | claude -p "analyze this"
    .\checkup.ps1 -Save checkup-report.txt
#>

[CmdletBinding()]
param(
    [switch]$Report,
    [switch]$Clipboard,
    [switch]$Analyze,
    [string]$Ai,
    [string]$Model,
    [string]$Save
)

# ── Globals ──────────────────────────────────────────────────────────────────

$script:IssueCount  = 0
$script:WarnCount   = 0
$script:Issues      = [System.Collections.Generic.List[string]]::new()
$script:ReportLines = [System.Collections.Generic.List[string]]::new()
$script:Quiet       = $Report -or $Clipboard -or $Analyze -or $Save

# ── Helpers ──────────────────────────────────────────────────────────────────

function Write-Header {
    param([string]$Title)
    $line = "─" * 60

    $script:ReportLines.Add("")
    $script:ReportLines.Add("═══ $Title $("═" * (55 - $Title.Length))")

    if ($script:Quiet) { return }
    Write-Host ""
    Write-Host "  ┌$line┐" -ForegroundColor DarkCyan
    Write-Host "  │ " -ForegroundColor DarkCyan -NoNewline
    Write-Host "$($Title.PadRight(58))" -ForegroundColor Cyan -NoNewline
    Write-Host " │" -ForegroundColor DarkCyan
    Write-Host "  └$line┘" -ForegroundColor DarkCyan
}

function Write-Metric {
    param(
        [string]$Label,
        [string]$Value,
        [string]$Status = "ok",      # ok, warn, bad, info
        [string]$Note = ""
    )

    # Build plain-text line
    $icon = switch ($Status) {
        "ok"   { "[OK]  " }
        "warn" { "[WARN]" }
        "bad"  { "[BAD] " }
        "info" { "[INFO]" }
        default { "[INFO]" }
    }
    $plainLine = "  $icon $($Label.PadRight(34)) $Value"
    if ($Note) { $plainLine += "  << $Note" }
    $script:ReportLines.Add($plainLine)

    # Track issues
    if ($Status -eq "bad") {
        $script:IssueCount++
        $script:Issues.Add("[$Label] $Value $(if ($Note) { "($Note)" })")
    } elseif ($Status -eq "warn") {
        $script:WarnCount++
        $script:Issues.Add("[WARN: $Label] $Value $(if ($Note) { "($Note)" })")
    }

    if ($script:Quiet) { return }

    # Fancy output
    $fancyIcon = switch ($Status) {
        "ok"   { "✓" }
        "warn" { "▲" }
        "bad"  { "✗" }
        "info" { "●" }
        default { "●" }
    }
    $iconColor = switch ($Status) {
        "ok"   { "Green" }
        "warn" { "Yellow" }
        "bad"  { "Red" }
        "info" { "DarkGray" }
        default { "DarkGray" }
    }

    Write-Host "    $fancyIcon " -ForegroundColor $iconColor -NoNewline
    $padLen = [Math]::Max(32, $Label.Length + 2)
    Write-Host "$($Label.PadRight($padLen))" -ForegroundColor White -NoNewline
    Write-Host "$Value" -ForegroundColor Gray -NoNewline
    if ($Note) {
        Write-Host "  $Note" -ForegroundColor DarkYellow
    } else {
        Write-Host ""
    }
}

function Write-SubHeader {
    param([string]$Title)
    $script:ReportLines.Add("  --- $Title ---")
    if ($script:Quiet) { return }
    Write-Host ""
    Write-Host "    $Title" -ForegroundColor DarkYellow
    Write-Host "    $("─" * $Title.Length)" -ForegroundColor DarkGray
}

function Get-FriendlySize {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "{0:N1} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N1} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N0} MB" -f ($Bytes / 1MB) }
    return "{0:N0} KB" -f ($Bytes / 1KB)
}

# ── Banner ───────────────────────────────────────────────────────────────────

$scanStart = Get-Date

if (-not $script:Quiet) {
    try { Clear-Host } catch {}
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║                                                              ║" -ForegroundColor Cyan
    Write-Host "  ║" -ForegroundColor Cyan -NoNewline
    Write-Host "              PC CHECKUP  ·  Health & Performance            " -ForegroundColor White -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host "  ║" -ForegroundColor Cyan -NoNewline
    Write-Host "              ─────────────────────────────────              " -ForegroundColor DarkGray -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host "  ║" -ForegroundColor Cyan -NoNewline
    $dateStr = (Get-Date -Format 'yyyy-MM-dd HH:mm').PadLeft(38).PadRight(62)
    Write-Host $dateStr -ForegroundColor DarkGray -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host "  ║                                                              ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
}

$script:ReportLines.Add("PC CHECKUP - Health & Performance Report")
$script:ReportLines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    $script:ReportLines.Add("Note: Running without admin. Some checks may be limited.")
    if (-not $script:Quiet) {
        Write-Host ""
        Write-Host "    ● Running without admin. Some checks may be limited." -ForegroundColor DarkGray
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# CHECKS
# ══════════════════════════════════════════════════════════════════════════════

# ── 1. System Overview ──────────────────────────────────────────────────────

Write-Header "SYSTEM OVERVIEW"

$cs  = Get-CimInstance Win32_ComputerSystem
$os  = Get-CimInstance Win32_OperatingSystem
$cpu = Get-CimInstance Win32_Processor
$gpu = Get-CimInstance Win32_VideoController

Write-Metric "Machine" "$($cs.Manufacturer) $($cs.Model)"
Write-Metric "OS" "$($os.Caption) (Build $($os.BuildNumber))"
Write-Metric "CPU" "$($cpu.Name.Trim())" -Status "info"
Write-Metric "Cores / Threads" "$($cpu.NumberOfCores) / $($cpu.NumberOfLogicalProcessors)" -Status "info"

# Win32_VideoController.AdapterRAM is a 32-bit value (caps at 4 GB); the
# driver registry key stores the real size as a 64-bit qwMemorySize.
$gpuRegSizes = @{}
Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0*" -ErrorAction SilentlyContinue |
    Where-Object { $_.'HardwareInformation.qwMemorySize' -gt 0 } |
    ForEach-Object { $gpuRegSizes[$_.DriverDesc] = $_.'HardwareInformation.qwMemorySize' }

foreach ($g in $gpu) {
    $vramBytes = if ($gpuRegSizes.ContainsKey($g.Name)) { $gpuRegSizes[$g.Name] }
                 elseif ($g.AdapterRAM -gt 0) { $g.AdapterRAM } else { 0 }
    $vramStr = if ($vramBytes -gt 0) { " ($(Get-FriendlySize $vramBytes))" } else { "" }
    Write-Metric "GPU" "$($g.Name)$vramStr" -Status "info"
}

$disks = Get-PhysicalDisk
foreach ($disk in $disks) {
    $diskType = if ($disk.MediaType -eq "SSD") { "ok" } elseif ($disk.MediaType -eq "HDD") { "bad" } else { "info" }
    $diskNote = if ($disk.MediaType -eq "HDD") { "HDD is a major bottleneck, consider SSD" } else { "" }
    Write-Metric "Disk" "$(Get-FriendlySize $disk.Size) $($disk.MediaType) ($($disk.HealthStatus))" -Status $diskType -Note $diskNote
}

# BIOS / Firmware
$bios = Get-CimInstance Win32_BIOS
Write-Metric "BIOS" "$($bios.SMBIOSBIOSVersion) ($($bios.ReleaseDate.ToString('yyyy-MM-dd')))" -Status "info"

# Display
try {
    $monitors = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -ErrorAction SilentlyContinue
    $videoSettings = Get-CimInstance Win32_VideoController | Select-Object -First 1
    if ($videoSettings.CurrentHorizontalResolution) {
        $refresh = $videoSettings.CurrentRefreshRate
        $refreshStatus = if ($refresh -gt 60) { "ok" } else { "info" }
        Write-Metric "Display" "$($videoSettings.CurrentHorizontalResolution)x$($videoSettings.CurrentVerticalResolution) @ ${refresh}Hz" -Status $refreshStatus
    }
} catch {}

# Uptime
$uptime = (Get-Date) - $os.LastBootUpTime
$uptimeStr = if ($uptime.Days -gt 0) { "$($uptime.Days)d $($uptime.Hours)h $($uptime.Minutes)m" } else { "$($uptime.Hours)h $($uptime.Minutes)m" }
$uptimeStatus = if ($uptime.Days -ge 14) { "warn" } elseif ($uptime.Days -ge 7) { "info" } else { "ok" }
Write-Metric "Uptime" $uptimeStr -Status $uptimeStatus -Note $(if ($uptime.Days -ge 14) { "Consider rebooting for fresh state" })

# ── 2. Memory ────────────────────────────────────────────────────────────────

Write-Header "MEMORY"

$totalRAM = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
$freeRAM  = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
$usedRAM  = [math]::Round($totalRAM - $freeRAM, 1)
$usedPct  = [math]::Round(($usedRAM / $totalRAM) * 100, 0)

$ramStatus = if ($usedPct -ge 85) { "bad" } elseif ($usedPct -ge 70) { "warn" } else { "ok" }
$ramNote   = if ($usedPct -ge 85) { "Critical! Close apps before meetings" } elseif ($usedPct -ge 70) { "High idle usage" } else { "" }

Write-Metric "Total RAM" "${totalRAM} GB"
Write-Metric "Used" "${usedRAM} GB (${usedPct}%)" -Status $ramStatus -Note $ramNote
Write-Metric "Free" "${freeRAM} GB" -Status "info"

# Memory compression
$memComp = Get-Process "Memory Compression" -ErrorAction SilentlyContinue
if ($memComp) {
    $compMB = [math]::Round($memComp.WorkingSet64 / 1MB, 0)
    $compStatus = if ($compMB -ge 1000) { "warn" } else { "ok" }
    Write-Metric "Memory Compression" "${compMB} MB" -Status $compStatus -Note $(if ($compMB -ge 1000) { "System is under memory pressure" })
}

# Pagefile
$pagefile = Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue
if ($pagefile) {
    $pfUsePct = if ($pagefile.AllocatedBaseSize -gt 0) { [math]::Round(($pagefile.CurrentUsage / $pagefile.AllocatedBaseSize) * 100, 0) } else { 0 }
    $pfStatus = if ($pfUsePct -ge 80) { "warn" } else { "ok" }
    Write-Metric "Pagefile" "$($pagefile.CurrentUsage) MB / $($pagefile.AllocatedBaseSize) MB ($pfUsePct%)" -Status $pfStatus
}

# RAM configuration (single-channel is a major performance drain)
$dimms = @(Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue)
if ($dimms.Count -gt 0) {
    $runSpeed   = ($dimms | Select-Object -First 1).ConfiguredClockSpeed
    $ratedSpeed = ($dimms | Measure-Object Speed -Maximum).Maximum
    $dimmStatus = if ($dimms.Count -eq 1) { "warn" } else { "ok" }
    $dimmNote   = if ($dimms.Count -eq 1) { "Single stick = single channel, big performance hit" } else { "" }
    Write-Metric "RAM Config" "$($dimms.Count) stick(s) @ ${runSpeed} MHz" -Status $dimmStatus -Note $dimmNote
    if ($ratedSpeed -gt $runSpeed) {
        Write-Metric "RAM Speed" "Running ${runSpeed} MHz, rated ${ratedSpeed} MHz" -Status "warn" -Note "Enable XMP/EXPO in BIOS"
    }
}

Write-SubHeader "Top memory consumers"

$topProcs = Get-Process | Group-Object -Property Name |
    ForEach-Object {
        [PSCustomObject]@{
            Name   = $_.Name
            Count  = $_.Count
            RAM_MB = [math]::Round(($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
        }
    } | Sort-Object RAM_MB -Descending | Select-Object -First 10

foreach ($p in $topProcs) {
    $countStr = if ($p.Count -gt 1) { " (x$($p.Count))" } else { "" }
    $pStatus = if ($p.RAM_MB -ge 1000) { "warn" } else { "info" }
    Write-Metric "$($p.Name)$countStr" "$('{0:N0}' -f $p.RAM_MB) MB" -Status $pStatus
}

# ── 3. Battery ───────────────────────────────────────────────────────────────

Write-Header "BATTERY"

$battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
if ($battery) {
    $charge = $battery.EstimatedChargeRemaining
    $statusText = switch ($battery.BatteryStatus) {
        1 { "Discharging" }; 2 { "AC Power" }; 3 { "Fully Charged" }
        4 { "Low" }; 5 { "Critical" }; default { "Unknown" }
    }
    Write-Metric "Status" "$statusText ($charge%)" -Status $(if ($charge -le 20 -and $battery.BatteryStatus -eq 1) { "warn" } else { "ok" })

    $reportPath = "$env:TEMP\pc-checkup-batt.xml"
    $null = powercfg /batteryreport /output $reportPath /xml 2>$null
    if (Test-Path $reportPath) {
        try {
            [xml]$battReport = Get-Content $reportPath
            $designCap     = [int]$battReport.BatteryReport.Batteries.Battery.DesignCapacity
            $fullChargeCap = [int]$battReport.BatteryReport.Batteries.Battery.FullChargeCapacity
            if ($designCap -gt 0) {
                $healthPct  = [math]::Round(($fullChargeCap / $designCap) * 100, 1)
                $battStatus = if ($healthPct -le 60) { "bad" } elseif ($healthPct -le 80) { "warn" } else { "ok" }
                $battNote   = if ($healthPct -le 60) { "Battery is degraded, consider replacement" } elseif ($healthPct -le 80) { "Some wear, monitor over time" } else { "" }
                Write-Metric "Health" "${healthPct}% of design capacity" -Status $battStatus -Note $battNote
                Write-Metric "Design / Actual" "$([math]::Round($designCap/1000,1)) Wh / $([math]::Round($fullChargeCap/1000,1)) Wh" -Status "info"
            }
        } catch {}
        Remove-Item $reportPath -ErrorAction SilentlyContinue
    }
} else {
    Write-Metric "Battery" "Not detected (desktop?)" -Status "info"
}

# ── 4. Thermals & CPU ────────────────────────────────────────────────────────

Write-Header "THERMALS & CPU"

# Try to get CPU temperature
$tempFound = $false
try {
    $thermalZones = Get-CimInstance MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" -ErrorAction Stop
    foreach ($tz in $thermalZones) {
        $tempC = [math]::Round(($tz.CurrentTemperature - 2732) / 10, 1)
        if ($tempC -gt 0 -and $tempC -lt 120) {
            $tempStatus = if ($tempC -ge 90) { "bad" } elseif ($tempC -ge 75) { "warn" } else { "ok" }
            $tempNote = if ($tempC -ge 90) { "Throttling likely!" } elseif ($tempC -ge 75) { "Running warm" } else { "" }
            Write-Metric "Temperature" "${tempC}°C" -Status $tempStatus -Note $tempNote
            $tempFound = $true
        }
    }
} catch {}

# Fallback: Thermal Zone perf counter (sometimes readable without admin)
if (-not $tempFound) {
    try {
        $tzSamples = (Get-Counter '\Thermal Zone Information(*)\Temperature' -ErrorAction Stop).CounterSamples
        foreach ($s in $tzSamples) {
            $tempC = [math]::Round($s.CookedValue - 273.15, 1)
            if ($tempC -gt 0 -and $tempC -lt 120) {
                $tempStatus = if ($tempC -ge 90) { "bad" } elseif ($tempC -ge 75) { "warn" } else { "ok" }
                Write-Metric "Temperature ($($s.InstanceName))" "${tempC}°C" -Status $tempStatus
                $tempFound = $true
            }
        }
    } catch {}
}

# Fallback: LibreHardwareMonitor / OpenHardwareMonitor WMI (app must be running)
if (-not $tempFound) {
    foreach ($ns in @("root\LibreHardwareMonitor", "root\OpenHardwareMonitor")) {
        try {
            $sensors = @(Get-CimInstance -Namespace $ns -ClassName Sensor -ErrorAction Stop |
                Where-Object { $_.SensorType -eq 'Temperature' -and $_.Name -match 'CPU|Core|Tctl|Tdie|Package' })
            foreach ($s in $sensors | Sort-Object Value -Descending | Select-Object -First 1) {
                $tempC = [math]::Round($s.Value, 1)
                if ($tempC -gt 0 -and $tempC -lt 120) {
                    $tempStatus = if ($tempC -ge 90) { "bad" } elseif ($tempC -ge 75) { "warn" } else { "ok" }
                    Write-Metric "Temperature ($($s.Name))" "${tempC}°C" -Status $tempStatus
                    $tempFound = $true
                }
            }
            if ($tempFound) { break }
        } catch {}
    }
}

# Fallback: HWiNFO sensors published to registry (Settings > "Enable VSB")
if (-not $tempFound) {
    $vsb = Get-ItemProperty "HKCU:\SOFTWARE\HWiNFO64\VSB" -ErrorAction SilentlyContinue
    if ($vsb) {
        $props = $vsb.PSObject.Properties | Where-Object { $_.Name -match '^Label\d+$' -and $_.Value -match 'CPU|Tctl|Tdie' }
        foreach ($p in $props | Select-Object -First 1) {
            $idx = $p.Name -replace 'Label', ''
            $valStr = $vsb."Value$idx"
            if ($valStr -match '([\d,.]+)') {
                $tempC = [double]($Matches[1] -replace ',', '.')
                if ($tempC -gt 0 -and $tempC -lt 120) {
                    $tempStatus = if ($tempC -ge 90) { "bad" } elseif ($tempC -ge 75) { "warn" } else { "ok" }
                    Write-Metric "Temperature ($($p.Value))" "${tempC}°C" -Status $tempStatus
                    $tempFound = $true
                }
            }
        }
    }
}

if (-not $tempFound) {
    Write-Metric "Temperature" "Not exposed by this board's firmware" -Status "info" -Note "Run LibreHardwareMonitor or HWiNFO to provide it (admin alone won't help)"
}

# Thermal throttling — multi-signal check
# Signal 1: event log. Provider MUST be pinned: Id 37 is also used by
# Time-Service (NTP sync) and others; unpinned Id filters count unrelated events.
# Id 55 from Kernel-Processor-Power is a per-core informational capabilities dump
# at every boot on many builds, so only Warning/Error level (Level <= 3) counts.
$throttleEvents = $null   # $null = query failed; empty array = no events
try {
    $throttleEvents = @(Get-WinEvent -FilterHashtable @{
        LogName      = 'System'
        ProviderName = 'Microsoft-Windows-Kernel-Processor-Power'
        Id           = @(37, 55)
        StartTime    = (Get-Date).AddDays(-7)
    } -ErrorAction Stop) | Where-Object { $_.Level -ge 1 -and $_.Level -le 3 }
    $throttleEvents = @($throttleEvents)
} catch {
    if ($_.Exception.Message -match 'No events were found') { $throttleEvents = @() }
}

# Signal 2: live clock ratio. Win32_Processor.CurrentClockSpeed is often pinned
# to base clock on modern CPUs, so prefer the perf counter (boost can exceed 100%).
$clockPct = $null
try {
    $clockPct = [math]::Round((Get-Counter '\Processor Information(_Total)\% Processor Performance' -ErrorAction Stop).CounterSamples[0].CookedValue, 0)
} catch {
    # counter name is localized on non-English Windows; fall back to WMI ratio
    try {
        $cpuNow = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
        if ($cpuNow.MaxClockSpeed -gt 0) {
            $clockPct = [math]::Round(($cpuNow.CurrentClockSpeed / $cpuNow.MaxClockSpeed) * 100, 0)
        }
    } catch {}
}

if ($null -eq $throttleEvents) {
    Write-Metric "Thermal Throttling" "Could not check event log" -Status "info"
} elseif ($throttleEvents.Count -gt 0) {
    $liveNote = if ($null -ne $clockPct -and $clockPct -lt 80) {
        "CPU throttled now (clock at ${clockPct}% of nominal) and $($throttleEvents.Count) firmware-limit event(s) in 7 days"
    } else {
        "$($throttleEvents.Count) firmware-limit event(s) in 7 days; not throttled right now"
    }
    $liveStatus = if ($null -ne $clockPct -and $clockPct -lt 80) { "bad" } else { "warn" }
    Write-Metric "Thermal Throttling" "$($throttleEvents.Count) event(s) in last 7 days" -Status $liveStatus -Note $liveNote
} else {
    Write-Metric "Thermal Throttling" "No events in last 7 days" -Status "ok"
}
if ($null -ne $clockPct) {
    Write-Metric "CPU Clock (now)" "${clockPct}% of nominal" -Status $(if ($clockPct -lt 50 -and $throttleEvents.Count -gt 0) { "warn" } else { "info" })
}

# CPU current load
try {
    $cpuLoad = (Get-CimInstance Win32_Processor).LoadPercentage
    $loadStatus = if ($cpuLoad -ge 80) { "bad" } elseif ($cpuLoad -ge 50) { "warn" } else { "ok" }
    Write-Metric "CPU Load (now)" "${cpuLoad}%" -Status $loadStatus
} catch {}

# ── 5. Disk Performance ─────────────────────────────────────────────────────

Write-Header "DISK PERFORMANCE"

foreach ($disk in $disks) {
    Write-Metric "Disk" "$(Get-FriendlySize $disk.Size) $($disk.MediaType)" -Status "info"

    # Temperature and wear if available (query this disk only, not the whole set)
    try {
        $rel = $disk | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue
        if ($rel.Temperature -gt 0) {
            $dtStatus = if ($rel.Temperature -ge 60) { "bad" } elseif ($rel.Temperature -ge 50) { "warn" } else { "ok" }
            Write-Metric "  Temperature" "$($rel.Temperature)°C" -Status $dtStatus
        }
        if ($null -ne $rel.Wear -and $rel.Wear -gt 0) {
            $wearStatus = if ($rel.Wear -ge 80) { "bad" } elseif ($rel.Wear -ge 50) { "warn" } else { "ok" }
            Write-Metric "  Wear" "$($rel.Wear)% used" -Status $wearStatus -Note $(if ($rel.Wear -ge 80) { "SSD nearing end of life" })
        }
    } catch {}
}

# TRIM (SSD performance degrades badly without it)
$trimOut = (fsutil behavior query DisableDeleteNotify 2>$null) -join "`n"
if ($trimOut -match 'NTFS DisableDeleteNotify = (\d)') {
    $trimOn = $Matches[1] -eq '0'
    Write-Metric "TRIM (NTFS)" $(if ($trimOn) { "Enabled" } else { "DISABLED" }) -Status $(if ($trimOn) { "ok" } else { "bad" }) -Note $(if (-not $trimOn) { "Enable: fsutil behavior set DisableDeleteNotify 0" })
}

# Disk usage per volume
$volumes = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
foreach ($vol in $volumes) {
    $totalGB = [math]::Round($vol.Size / 1GB, 0)
    $freeGB  = [math]::Round($vol.FreeSpace / 1GB, 0)
    $usedGB  = $totalGB - $freeGB
    $usePct  = if ($totalGB -gt 0) { [math]::Round(($usedGB / $totalGB) * 100, 0) } else { 0 }
    $volStatus = if ($usePct -ge 95) { "bad" } elseif ($usePct -ge 85) { "warn" } else { "ok" }
    $volNote = if ($usePct -ge 95) { "Nearly full! Free up space" } elseif ($usePct -ge 85) { "Getting full" } else { "" }
    Write-Metric "Volume $($vol.DeviceID)" "${usedGB} GB / ${totalGB} GB (${usePct}%)" -Status $volStatus -Note $volNote
}

# ── 6. Power Plan ────────────────────────────────────────────────────────────

Write-Header "POWER PLAN"

$activePlan = powercfg /getactivescheme
$planName   = if ($activePlan -match '\((.+)\)') { $Matches[1] } else { "Unknown" }
$planStatus = if ($planName -match "Balanced|Power saver") { "warn" } else { "ok" }
$planNote   = if ($planName -match "Balanced") { "Switch to Best Performance for calls" } elseif ($planName -match "Power saver") { "Will throttle CPU heavily" } else { "" }
Write-Metric "Active Plan" $planName -Status $planStatus -Note $planNote

# powercfg returns an array of lines; join before matching, otherwise the
# multi-line regex never matches any single line and the check silently skips.
$procPower = (powercfg /query SCHEME_CURRENT SUB_PROCESSOR 2>$null) -join "`n"
$minProc = if ($procPower -match "Minimum processor state[\s\S]*?Current AC Power Setting Index:\s+0x([0-9a-fA-F]+)") { [int]("0x$($Matches[1])") } else { -1 }
$maxProc = if ($procPower -match "Maximum processor state[\s\S]*?Current AC Power Setting Index:\s+0x([0-9a-fA-F]+)") { [int]("0x$($Matches[1])") } else { -1 }
if ($minProc -ge 0) { Write-Metric "CPU Min State (AC)" "${minProc}%" -Status "info" }
if ($maxProc -ge 0) {
    $maxStatus = if ($maxProc -lt 100) { "warn" } else { "ok" }
    Write-Metric "CPU Max State (AC)" "${maxProc}%" -Status $maxStatus -Note $(if ($maxProc -lt 100) { "CPU is being capped!" })
}

# Fast Startup (shutdown becomes hibernate-lite; stale state survives "shutdown")
$hib = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name HiberbootEnabled -ErrorAction SilentlyContinue
if ($null -ne $hib) {
    Write-Metric "Fast Startup" $(if ($hib.HiberbootEnabled -eq 1) { "Enabled" } else { "Disabled" }) -Status "info" -Note $(if ($hib.HiberbootEnabled -eq 1) { "Use Restart (not Shutdown) for a true fresh boot" })
}

# ── 7. Security Overhead ────────────────────────────────────────────────────

Write-Header "SECURITY & VIRTUALIZATION"

$dg = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard -ErrorAction SilentlyContinue
if ($dg) {
    $vbsStatus = $dg.VirtualizationBasedSecurityStatus
    $services  = $dg.SecurityServicesRunning
    $vbsOn  = $vbsStatus -eq 2
    $hvciOn = $services -contains 2 -or $services -contains 3

    Write-Metric "VBS" $(if ($vbsOn) { "ENABLED" } else { "Disabled" }) `
        -Status $(if ($vbsOn) { "warn" } else { "ok" }) `
        -Note $(if ($vbsOn) { "5-15% CPU overhead (Virtualization Based Security)" })

    Write-Metric "HVCI (Memory Integrity)" $(if ($hvciOn) { "ENABLED" } else { "Disabled" }) `
        -Status $(if ($hvciOn) { "warn" } else { "ok" }) `
        -Note $(if ($hvciOn) { "Core Isolation > Memory Integrity to disable" })

    if ($services -contains 4) {
        Write-Metric "System Guard Secure Launch" "ENABLED" -Status "info"
    }
} else {
    Write-Metric "VBS / HVCI" "Could not query (needs admin?)" -Status "info"
}

$hyperv = Get-Service vmcompute -ErrorAction SilentlyContinue
$hypervOn = $hyperv -and $hyperv.Status -eq "Running"
Write-Metric "Hyper-V" $(if ($hypervOn) { "Running" } else { "Not running" }) -Status $(if ($hypervOn) { "info" } else { "ok" })

# ── 8. Windows Defender ──────────────────────────────────────────────────────

Write-Header "WINDOWS DEFENDER"

$cpuCap = $null
try {
    $mpStatus = Get-MpComputerStatus -ErrorAction Stop
    $mpPref   = Get-MpPreference -ErrorAction Stop

    Write-Metric "Real-time Protection" $(if ($mpStatus.RealTimeProtectionEnabled) { "ON" } else { "OFF" }) -Status "info"

    $cpuCap = $mpPref.ScanAvgCPULoadFactor
    $cpuCapStatus = if ($cpuCap -ge 50) { "warn" } else { "ok" }
    Write-Metric "Scan CPU Limit" "${cpuCap}%" -Status $cpuCapStatus -Note $(if ($cpuCap -ge 50) { "Set to 20%: Set-MpPreference -ScanAvgCPULoadFactor 20" })

    $exclusions = try { $mpPref.ExclusionPath } catch { $null }
    if ($exclusions -and $exclusions.Count -gt 0) {
        Write-Metric "Exclusion Paths" "$($exclusions.Count) configured" -Status "ok"
    } else {
        Write-Metric "Exclusion Paths" "None (add dev folders for speed)" -Status "warn" -Note "e.g. projects, node_modules"
    }

    Write-Metric "Last Quick Scan" "$($mpStatus.QuickScanAge) day(s) ago" -Status "info"
    Write-Metric "Signature Age" "$($mpStatus.AntivirusSignatureAge) day(s)" -Status $(if ($mpStatus.AntivirusSignatureAge -ge 3) { "warn" } else { "ok" })
} catch {
    Write-Metric "Defender" "Could not query (permission denied?)" -Status "info"
}

# ── 9. WSL / Docker ─────────────────────────────────────────────────────────

Write-Header "WSL / DOCKER"

$vmmem = Get-Process vmmem -ErrorAction SilentlyContinue
if ($vmmem) {
    $vmmemMB = [math]::Round($vmmem.WorkingSet64 / 1MB, 0)
    $vmmemStatus = if ($vmmemMB -ge 4000) { "bad" } elseif ($vmmemMB -ge 2000) { "warn" } else { "ok" }
    Write-Metric "WSL2 VM (vmmem)" "$('{0:N0}' -f $vmmemMB) MB" -Status $vmmemStatus -Note $(if ($vmmemMB -ge 4000) { "WSL is eating RAM!" })
} else {
    Write-Metric "WSL2 VM (vmmem)" "Not running" -Status "ok"
}

$wslConfig = "$env:USERPROFILE\.wslconfig"
if (Test-Path $wslConfig) {
    $wslContent = Get-Content $wslConfig -Raw
    $memMatch = if ($wslContent -match 'memory\s*=\s*(\S+)') { $Matches[1] } else { "not set (default: 50% RAM)" }
    Write-Metric ".wslconfig memory" $memMatch -Status $(if ($wslContent -match 'memory\s*=') { "ok" } else { "warn" })
} else {
    Write-Metric ".wslconfig" "Missing! WSL can use up to 50% RAM" -Status "warn" -Note "Create to cap WSL memory"
}

$docker = Get-Process "Docker Desktop" -ErrorAction SilentlyContinue
Write-Metric "Docker Desktop" $(if ($docker) { "Running" } else { "Not running" }) -Status $(if ($docker) { "info" } else { "ok" })

# ── 10. Startup Programs ────────────────────────────────────────────────────

Write-Header "STARTUP PROGRAMS"

$startupItems = @()

# Items disabled in Task Manager stay in Run/Startup but are marked disabled in
# StartupApproved (first byte odd = disabled). Don't count those as active.
$disabledStartup = @{}
foreach ($approvedKey in @(
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder"
)) {
    $approved = Get-ItemProperty $approvedKey -ErrorAction SilentlyContinue
    if ($approved) {
        $approved.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' -and $_.Value -is [byte[]] -and $_.Value.Length -gt 0 } | ForEach-Object {
            if (($_.Value[0] -band 1) -eq 1) { $disabledStartup[$_.Name] = $true }
        }
    }
}

$hkcuRun = Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -ErrorAction SilentlyContinue
if ($hkcuRun) {
    $hkcuRun.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
        $startupItems += [PSCustomObject]@{ Name = $_.Name; Source = "Registry (User)"; Disabled = $disabledStartup.ContainsKey($_.Name) }
    }
}

$hklmRun = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -ErrorAction SilentlyContinue
if ($hklmRun) {
    $hklmRun.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
        $startupItems += [PSCustomObject]@{ Name = $_.Name; Source = "Registry (System)"; Disabled = $disabledStartup.ContainsKey($_.Name) }
    }
}

$startupFolder = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup")
if (Test-Path $startupFolder) {
    Get-ChildItem $startupFolder -File | ForEach-Object {
        $startupItems += [PSCustomObject]@{ Name = $_.BaseName; Source = "Startup Folder"; Disabled = $disabledStartup.ContainsKey($_.Name) }
    }
}

$heavyApps = @("Docker Desktop", "Discord", "Spotify", "Steam", "Teams", "Slack", "OneDrive", "Figma Agent", "ShareX")

foreach ($item in $startupItems | Sort-Object Name) {
    if ($item.Disabled) {
        Write-Metric $item.Name "$($item.Source) — disabled" -Status "info"
        continue
    }
    $isHeavy = $heavyApps | Where-Object { $item.Name -match [regex]::Escape($_) -or $item.Name -match ($_ -replace ' ','') }
    $status = if ($isHeavy) { "warn" } else { "info" }
    $note   = if ($isHeavy) { "Consider disabling, start manually" } else { "" }
    Write-Metric $item.Name $item.Source -Status $status -Note $note
}

$activeStartup = @($startupItems | Where-Object { -not $_.Disabled })
Write-Metric "Active startup items" "$($activeStartup.Count) ($($startupItems.Count) total)" -Status $(if ($activeStartup.Count -ge 8) { "warn" } else { "ok" })

# ── 11. Bloat & Background Processes ────────────────────────────────────────

Write-Header "BACKGROUND BLOAT"

$bloatChecks = @(
    @{ Pattern = "WidgetBoard|WidgetService";  Label = "Windows Widgets" }
    @{ Pattern = "PhoneExperienceHost";         Label = "Phone Link" }
    @{ Pattern = "msedgewebview2";              Label = "Edge WebView2" }
    @{ Pattern = "GameBar|gamebar";             Label = "Xbox Game Bar" }
    @{ Pattern = "Cortana";                     Label = "Cortana" }
    @{ Pattern = "OneDrive";                    Label = "OneDrive" }
    @{ Pattern = "YourPhone";                   Label = "Your Phone" }
)

foreach ($check in $bloatChecks) {
    $procs = Get-Process | Where-Object { $_.Name -match $check.Pattern }
    if ($procs) {
        $totalMB = [math]::Round(($procs | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
        $count = $procs.Count
        $status = if ($totalMB -ge 200) { "warn" } else { "info" }
        Write-Metric "$($check.Label) (x$count)" "$('{0:N0}' -f $totalMB) MB" -Status $status -Note $(if ($totalMB -ge 200) { "Disable to save resources" })
    }
}

$electronApps = @("Discord", "Slack", "Spotify", "Teams", "code", "claude", "figma", "notion", "obsidian")
$runningElectron = foreach ($app in $electronApps) {
    $p = Get-Process -Name $app -ErrorAction SilentlyContinue
    if ($p) { $app }
}
$electronCount = ($runningElectron | Select-Object -Unique).Count
$electronStatus = if ($electronCount -ge 4) { "warn" } else { "ok" }
Write-Metric "Electron/Chromium apps" "$electronCount ($($runningElectron -join ', '))" -Status $electronStatus -Note $(if ($electronCount -ge 4) { "Each runs own Chromium, huge GPU pressure" })

# ── 12. Network ──────────────────────────────────────────────────────────────

Write-Header "NETWORK"

$adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' -and $_.Name -notmatch 'vEthernet|Loopback' }
foreach ($a in $adapters) {
    Write-Metric $a.Name "$($a.InterfaceDescription) @ $($a.LinkSpeed)" -Status "info"
    try {
        $pm = Get-NetAdapterPowerManagement -Name $a.Name -ErrorAction SilentlyContinue
        if ($pm -and $pm.AllowComputerToTurnOffDevice -ne "Disabled") {
            Write-Metric "  Power Save" "Enabled (can drop WiFi)" -Status "warn" -Note "Disable in Device Manager"
        }
    } catch {}
}

$dns = (Get-DnsClientServerAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notmatch 'vEthernet|Loopback' } | Select-Object -First 1).ServerAddresses
if ($dns) {
    $dnsLabel = switch -Regex ($dns[0]) {
        "^1\.1\.1" { "Cloudflare ($($dns[0]))" }
        "^8\.8\."  { "Google ($($dns[0]))" }
        "^9\.9\."  { "Quad9 ($($dns[0]))" }
        "^192\.168|^10\.|^172\.(1[6-9]|2|3[01])" { "Router/ISP ($($dns[0]))" }
        default    { $dns[0] }
    }
    Write-Metric "DNS" $dnsLabel -Status "info"
}

# ── 13. Windows Update ───────────────────────────────────────────────────────

Write-Header "WINDOWS UPDATE"

try {
    $hotfixes = Get-HotFix | Sort-Object InstalledOn -Descending -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($hotfixes) {
        $daysSince = ((Get-Date) - $hotfixes.InstalledOn).Days
        $updateStatus = if ($daysSince -ge 60) { "warn" } else { "ok" }
        Write-Metric "Last Update" "$($hotfixes.InstalledOn.ToString('yyyy-MM-dd')) ($daysSince days ago)" -Status $updateStatus -Note $(if ($daysSince -ge 60) { "System may be missing patches" })
        Write-Metric "Latest KB" $hotfixes.HotFixID -Status "info"
    }
} catch {
    Write-Metric "Windows Update" "Could not query" -Status "info"
}

# Pending reboot check
$pendingReboot = $false
$cbsReboot = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction SilentlyContinue
$wuReboot  = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue
if ($cbsReboot -or $wuReboot) { $pendingReboot = $true }
Write-Metric "Pending Reboot" $(if ($pendingReboot) { "YES" } else { "No" }) -Status $(if ($pendingReboot) { "warn" } else { "ok" }) -Note $(if ($pendingReboot) { "Reboot to finish installing updates" })

# ── 14. Recent Crashes ───────────────────────────────────────────────────────

Write-Header "STABILITY"

try {
    # Provider pinned: Id 1000 in the Application log is also used by MsiInstaller,
    # LoadPerf and others; only 'Application Error' events are real crashes (and
    # only its schema has the faulting app name in Properties[0]).
    $appCrashes = @(Get-WinEvent -FilterHashtable @{
        LogName      = 'Application'
        ProviderName = 'Application Error'
        Id           = 1000
        StartTime    = (Get-Date).AddDays(-7)
    } -MaxEvents 10 -ErrorAction Stop)

    if ($appCrashes.Count -gt 0) {
        $crashedApps = $appCrashes | ForEach-Object {
            $_.Properties[0].Value
        } | Group-Object | Sort-Object Count -Descending | Select-Object -First 5

        Write-Metric "App crashes (7 days)" "$($appCrashes.Count) events" -Status $(if ($appCrashes.Count -ge 5) { "warn" } else { "info" })
        foreach ($ca in $crashedApps) {
            Write-Metric "  $($ca.Name)" "$($ca.Count) crash(es)" -Status "warn"
        }
    } else {
        Write-Metric "App crashes (7 days)" "None" -Status "ok"
    }
} catch {
    if ($_.Exception.Message -match 'No events were found') {
        Write-Metric "App crashes (7 days)" "None" -Status "ok"
    } else {
        Write-Metric "App crashes" "Could not check event log" -Status "info"
    }
}

# BSOD check
try {
    # Provider pinned: Id 1001 in the System log is also used by DHCP-Client,
    # Windows Error Reporting subcomponents, etc. Real BugCheck events come from
    # Microsoft-Windows-WER-SystemErrorReporting (shown as source "BugCheck").
    $bsods = @(Get-WinEvent -FilterHashtable @{
        LogName      = 'System'
        ProviderName = 'Microsoft-Windows-WER-SystemErrorReporting'
        Id           = 1001
        StartTime    = (Get-Date).AddDays(-30)
    } -MaxEvents 5 -ErrorAction Stop)

    if ($bsods.Count -gt 0) {
        Write-Metric "Blue Screens (30 days)" "$($bsods.Count) BSOD(s)" -Status "bad" -Note "Check minidump for details"
    } else {
        Write-Metric "Blue Screens (30 days)" "None" -Status "ok"
    }
} catch {
    if ($_.Exception.Message -match 'No events were found') {
        Write-Metric "Blue Screens (30 days)" "None" -Status "ok"
    } else {
        Write-Metric "Blue Screens" "Could not check" -Status "info"
    }
}

# ── 15. Misc Services ───────────────────────────────────────────────────────

Write-Header "MISC SERVICES"

$wsearch = Get-Service WSearch -ErrorAction SilentlyContinue
if ($wsearch -and $wsearch.Status -eq "Running") {
    Write-Metric "Windows Search Indexer" "Running" -Status "info" -Note "Can spike CPU/disk during indexing"
}

$anr = Get-Process AMDNoiseSuppression -ErrorAction SilentlyContinue
if ($anr) {
    Write-Metric "AMD Noise Suppression" "Running" -Status "warn" -Note "Duplicate if meeting app has its own"
}

# GPU driver age
foreach ($g in $gpu) {
    if ($g.DriverDate) {
        $driverAge = ((Get-Date) - $g.DriverDate).Days
        $drvStatus = if ($driverAge -ge 365) { "warn" } elseif ($driverAge -ge 180) { "info" } else { "ok" }
        Write-Metric "GPU Driver ($($g.Name))" "$($g.DriverVersion) ($driverAge days old)" -Status $drvStatus -Note $(if ($driverAge -ge 365) { "Consider updating GPU drivers" })
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ══════════════════════════════════════════════════════════════════════════════

$scanDuration = [math]::Round(((Get-Date) - $scanStart).TotalSeconds, 1)

$script:ReportLines.Add("")
$script:ReportLines.Add("═══ SUMMARY ═══════════════════════════════════════════════")
$script:ReportLines.Add("  Issues: $($script:IssueCount)  |  Warnings: $($script:WarnCount)  |  Scan time: ${scanDuration}s")

if ($script:Issues.Count -gt 0) {
    $script:ReportLines.Add("")
    $script:ReportLines.Add("  Findings:")
    foreach ($issue in $script:Issues) {
        $script:ReportLines.Add("    - $issue")
    }
}

# Build tips
$tips = @()

if ($dg -and $dg.VirtualizationBasedSecurityStatus -eq 2) {
    $tips += "Turn off Memory Integrity (Core Isolation > Memory Integrity) for 5-15% CPU gain"
}
if ($usedPct -ge 70) {
    $tips += "Close unused apps before meetings to free RAM (currently ${usedPct}% used at idle)"
}
if (-not (Test-Path "$env:USERPROFILE\.wslconfig")) {
    $tips += "Create ~/.wslconfig to cap WSL2 memory: [wsl2] memory=4GB"
}
if ($cpuCap -and $cpuCap -ge 50) {
    $tips += "Lower Defender scan CPU limit: Set-MpPreference -ScanAvgCPULoadFactor 20"
}
if ($electronCount -ge 3) {
    $tips += "Disable hardware acceleration in Electron apps (Discord, Slack, Spotify Settings)"
}
if ($planName -match "Balanced") {
    $tips += "Switch to Best Performance power mode during video calls"
}
if ($activeStartup.Count -ge 8) {
    $tips += "Reduce startup programs ($($activeStartup.Count) currently auto-start). Open Task Manager > Startup"
}
$bloatProcs = Get-Process | Where-Object { $_.Name -match "WidgetBoard|PhoneExperienceHost" }
if ($bloatProcs) {
    $tips += "Disable Windows Widgets (Taskbar settings) and Phone Link"
}
if ($uptime.Days -ge 14) {
    $tips += "Reboot your machine (uptime: $($uptime.Days) days)"
}

if ($tips.Count -eq 0) {
    $tips += "Looking good! No major issues found."
}

$script:ReportLines.Add("")
$script:ReportLines.Add("  Recommendations:")
for ($i = 0; $i -lt $tips.Count; $i++) {
    $script:ReportLines.Add("    $($i+1). $($tips[$i])")
}

# Fancy summary (console only)
if (-not $script:Quiet) {
    Write-Host ""
    $summaryColor = if ($script:IssueCount -gt 0) { "Red" } elseif ($script:WarnCount -gt 0) { "Yellow" } else { "Green" }
    Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor $summaryColor
    Write-Host "  ║" -ForegroundColor $summaryColor -NoNewline
    $summaryText = "  SUMMARY: $($script:IssueCount) issues, $($script:WarnCount) warnings (${scanDuration}s)"
    Write-Host "$($summaryText.PadRight(62))" -ForegroundColor $summaryColor -NoNewline
    Write-Host "║" -ForegroundColor $summaryColor
    Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor $summaryColor

    if ($script:Issues.Count -gt 0) {
        Write-Host ""
        Write-Host "  Findings:" -ForegroundColor White
        foreach ($issue in $script:Issues) {
            $color = if ($issue.StartsWith("[WARN")) { "Yellow" } else { "Red" }
            Write-Host "    · $issue" -ForegroundColor $color
        }
    }

    Write-Host ""
    Write-Host "  ┌────────────────────────────────────────────────────────────────┐" -ForegroundColor Magenta
    Write-Host "  │ " -ForegroundColor Magenta -NoNewline
    Write-Host "QUICK WINS                                                    " -ForegroundColor White -NoNewline
    Write-Host " │" -ForegroundColor Magenta
    Write-Host "  ├────────────────────────────────────────────────────────────────┤" -ForegroundColor Magenta

    for ($i = 0; $i -lt $tips.Count; $i++) {
        Write-Host "  │  " -ForegroundColor Magenta -NoNewline
        Write-Host "$($i+1). " -ForegroundColor Cyan -NoNewline
        $tipText = $tips[$i]
        $padded = $tipText.PadRight(60)
        if ($padded.Length -gt 60) { $padded = $tipText.Substring(0, 57) + "..." }
        Write-Host $padded -ForegroundColor Gray -NoNewline
        Write-Host " │" -ForegroundColor Magenta
    }

    Write-Host "  └────────────────────────────────────────────────────────────────┘" -ForegroundColor Magenta
}

# ══════════════════════════════════════════════════════════════════════════════
# AI INTEGRATION
# ══════════════════════════════════════════════════════════════════════════════

# Detect available AI CLIs
$aiTools = [ordered]@{}
$aiChecks = @(
    @{ Name = "claude";   Cmd = "claude";         Desc = "Claude Code" }
    @{ Name = "codex";    Cmd = "codex";           Desc = "OpenAI Codex CLI" }
    @{ Name = "copilot";  Cmd = "github-copilot";  Desc = "GitHub Copilot CLI" }
    @{ Name = "ollama";   Cmd = "ollama";           Desc = "Ollama (local)" }
    @{ Name = "aichat";   Cmd = "aichat";           Desc = "aichat" }
    @{ Name = "sgpt";     Cmd = "sgpt";             Desc = "ShellGPT" }
    @{ Name = "fabric";   Cmd = "fabric";           Desc = "Fabric" }
    @{ Name = "mods";     Cmd = "mods";             Desc = "Charmbracelet Mods" }
    @{ Name = "llm";      Cmd = "llm";              Desc = "LLM CLI (Simon Willison)" }
)

foreach ($aiEntry in $aiChecks) {
    $cmdName = $aiEntry["Cmd"]
    if ($cmdName) {
        $found = Get-Command $cmdName -ErrorAction SilentlyContinue
        if ($found) { $aiTools[$aiEntry["Name"]] = $aiEntry }
    }
}

# Build the full report text
$reportText = $script:ReportLines -join "`n"

$aiPrompt = @"
You are a Windows PC performance expert. Analyze this system health report and provide:

1. A brief overall health score (1-10) with one-line verdict
2. The top 3 most impactful issues to fix, ranked by performance gain
3. For each issue: what it is, why it matters, and the exact steps to fix it
4. Any issues that are fine and can be ignored
5. A "before your next video call" quick checklist

Be specific and actionable. Reference the actual values from the report.
Use short headers (## Header), numbered lists, and bullet points.
Keep it concise. No filler. No disclaimers.

--- SYSTEM REPORT ---
$reportText
"@

# Show detected AIs (always, unless in report mode)
if (-not $Report) {
    $aiListQuiet = $Clipboard -or $Analyze -or $Save
    if ($aiTools.Count -gt 0 -and -not $aiListQuiet) {
        Write-Host ""
        Write-Host "  ┌────────────────────────────────────────────────────────────────┐" -ForegroundColor Blue
        Write-Host "  │ " -ForegroundColor Blue -NoNewline
        Write-Host "AI TOOLS DETECTED                                             " -ForegroundColor White -NoNewline
        Write-Host " │" -ForegroundColor Blue
        Write-Host "  ├────────────────────────────────────────────────────────────────┤" -ForegroundColor Blue
        foreach ($key in $aiTools.Keys) {
            $tool = $aiTools[$key]
            Write-Host "  │  " -ForegroundColor Blue -NoNewline
            Write-Host "● $("$($tool['Desc'])".PadRight(20))" -ForegroundColor Green -NoNewline
            Write-Host "$("$($tool['Cmd'])".PadRight(40))" -ForegroundColor Gray -NoNewline
            Write-Host " │" -ForegroundColor Blue
        }
        Write-Host "  ├────────────────────────────────────────────────────────────────┤" -ForegroundColor Blue
        Write-Host "  │  " -ForegroundColor Blue -NoNewline
        Write-Host "Run: .\checkup.ps1 -Analyze" -ForegroundColor Cyan -NoNewline
        Write-Host "                                    │" -ForegroundColor Blue
        Write-Host "  │  " -ForegroundColor Blue -NoNewline
        Write-Host "Or:  .\checkup.ps1 -Analyze -Ai ollama -Model llama3" -ForegroundColor Cyan -NoNewline
        Write-Host "          │" -ForegroundColor Blue
        Write-Host "  │  " -ForegroundColor Blue -NoNewline
        Write-Host "Or:  .\checkup.ps1 -Clipboard" -ForegroundColor Cyan -NoNewline
        Write-Host " (copy for ChatGPT/Claude web)    │" -ForegroundColor DarkGray
        Write-Host "  └────────────────────────────────────────────────────────────────┘" -ForegroundColor Blue
    } elseif ($aiTools.Count -eq 0 -and -not $aiListQuiet) {
        Write-Host ""
        Write-Host "  ┌────────────────────────────────────────────────────────────────┐" -ForegroundColor DarkGray
        Write-Host "  │ " -ForegroundColor DarkGray -NoNewline
        Write-Host "No AI CLI detected. Use -Clipboard to copy for ChatGPT/Claude" -ForegroundColor Gray -NoNewline
        Write-Host " │" -ForegroundColor DarkGray
        Write-Host "  │ " -ForegroundColor DarkGray -NoNewline
        Write-Host "Install one: claude, codex, ollama, aichat, mods, llm, fabric" -ForegroundColor DarkGray -NoNewline
        Write-Host " │" -ForegroundColor DarkGray
        Write-Host "  └────────────────────────────────────────────────────────────────┘" -ForegroundColor DarkGray
    }
    Write-Host ""
}

# ── Handle output modes ─────────────────────────────────────────────────────

# -Report: dump plain text to stdout
if ($Report) {
    Write-Output $aiPrompt
    return
}

# -Save: write to file
if ($Save) {
    $aiPrompt | Out-File -FilePath $Save -Encoding UTF8
    Write-Host "  ✓ Report saved to: $Save" -ForegroundColor Green
    Write-Host ""
}

# -Clipboard: copy to clipboard
if ($Clipboard) {
    $aiPrompt | Set-Clipboard
    Write-Host "  ✓ Report copied to clipboard! Paste into ChatGPT, Claude, etc." -ForegroundColor Green
    Write-Host ""
    return
}

# -Analyze: pipe to local AI
if ($Analyze) {
    $selectedAi = $null

    if ($Ai) {
        # User specified a specific AI
        if ($aiTools.Contains($Ai)) {
            $selectedAi = $aiTools[$Ai]
        } else {
            # Try as raw command
            $rawCmd = Get-Command $Ai -ErrorAction SilentlyContinue
            if ($rawCmd) {
                $selectedAi = @{ Name = $Ai; Cmd = $Ai; Desc = $Ai }
            } else {
                Write-Host "  ✗ AI tool '$Ai' not found. Available: $($aiTools.Keys -join ', ')" -ForegroundColor Red
                return
            }
        }
    } else {
        # Auto-select best available
        $priority = @("claude", "codex", "copilot", "aichat", "mods", "fabric", "sgpt", "llm", "ollama")
        foreach ($p in $priority) {
            if ($aiTools.Contains($p)) {
                $selectedAi = $aiTools[$p]
                break
            }
        }
    }

    if (-not $selectedAi) {
        Write-Host "  ✗ No AI CLI found. Install one of: claude, codex, ollama, aichat, mods" -ForegroundColor Red
        Write-Host "    Or use: .\checkup.ps1 -Clipboard  (then paste into ChatGPT/Claude)" -ForegroundColor Gray
        return
    }

    Write-Host "  ● Analyzing with $($selectedAi['Desc'])..." -ForegroundColor Cyan

    $modelArg = ""

    switch ($selectedAi['Name']) {
        "claude" {
            # Claude Code: claude -p "prompt"
            $escapedPrompt = $aiPrompt -replace '"', '\"'
            $cmd = "claude -p `"$escapedPrompt`""
        }
        "codex" {
            # OpenAI Codex CLI
            $escapedPrompt = $aiPrompt -replace '"', '\"'
            $cmd = "codex -q `"$escapedPrompt`""
        }
        "ollama" {
            $ollamaModel = if ($Model) { $Model } else { "llama3.1" }
            $cmd = "echo `"$($aiPrompt -replace '"', '\"')`" | ollama run $ollamaModel"
        }
        "aichat" {
            $cmd = "echo `"$($aiPrompt -replace '"', '\"')`" | aichat"
        }
        "sgpt" {
            $cmd = "echo `"$($aiPrompt -replace '"', '\"')`" | sgpt"
        }
        "fabric" {
            $cmd = "echo `"$($aiPrompt -replace '"', '\"')`" | fabric"
        }
        "mods" {
            $cmd = "echo `"$($aiPrompt -replace '"', '\"')`" | mods"
        }
        "llm" {
            $modelFlag = if ($Model) { "-m $Model" } else { "" }
            $cmd = "echo `"$($aiPrompt -replace '"', '\"')`" | llm $modelFlag"
        }
        default {
            $cmd = "echo `"$($aiPrompt -replace '"', '\"')`" | $($selectedAi['Cmd'])"
        }
    }

    Write-Host ""

    # Use temp file for reliable piping (avoids shell escaping hell)
    $tempPrompt = "$env:TEMP\pc-checkup-prompt.txt"
    $aiPrompt | Out-File -FilePath $tempPrompt -Encoding UTF8

    try {
        switch ($selectedAi['Name']) {
            "claude" {
                $result = Get-Content $tempPrompt -Raw | & claude -p 2>&1
            }
            "codex" {
                $result = & codex -q (Get-Content $tempPrompt -Raw) 2>&1
            }
            "ollama" {
                $ollamaModel = if ($Model) { $Model } else { "llama3.1" }
                $result = Get-Content $tempPrompt -Raw | & ollama run $ollamaModel 2>&1
            }
            "aichat" {
                $result = Get-Content $tempPrompt -Raw | & aichat 2>&1
            }
            "mods" {
                $result = Get-Content $tempPrompt -Raw | & mods 2>&1
            }
            "llm" {
                $modelFlag = if ($Model) { @("-m", $Model) } else { @() }
                $result = Get-Content $tempPrompt -Raw | & llm @modelFlag 2>&1
            }
            default {
                $result = Get-Content $tempPrompt -Raw | & $selectedAi['Cmd'] 2>&1
            }
        }

        # Render markdown-ish output with basic terminal formatting
        $resultStr = ($result | Out-String).Trim()
        foreach ($line in $resultStr -split "`n") {
            $trimmed = $line.TrimEnd()
            if ($trimmed -match '^#{1,3}\s+(.+)') {
                # Headers
                Write-Host ""
                Write-Host "  $($Matches[1])" -ForegroundColor Cyan
                Write-Host "  $("─" * $Matches[1].Length)" -ForegroundColor DarkGray
            } elseif ($trimmed -match '^\*\*(.+?)\*\*(.*)') {
                # Bold start of line
                Write-Host "  $($Matches[1])" -ForegroundColor White -NoNewline
                Write-Host $Matches[2] -ForegroundColor Gray
            } elseif ($trimmed -match '^[-*]\s+(.+)') {
                # Bullet points
                Write-Host "  · $($Matches[1])" -ForegroundColor Gray
            } elseif ($trimmed -match '^\d+\.\s+(.+)') {
                # Numbered list
                $num = ($trimmed -split '\.')[0]
                Write-Host "  $num. " -ForegroundColor Cyan -NoNewline
                Write-Host ($trimmed -replace '^\d+\.\s+', '') -ForegroundColor Gray
            } elseif ($trimmed -match '^```') {
                # Code block delimiter (skip)
            } elseif ($trimmed -match '^>') {
                # Blockquote
                Write-Host "  │ $($trimmed -replace '^>\s*', '')" -ForegroundColor DarkYellow
            } elseif ($trimmed -eq '') {
                Write-Host ""
            } else {
                # Regular text, render inline **bold** and `code`
                $formatted = $trimmed -replace '\*\*(.+?)\*\*', '$1' -replace '`(.+?)`', '$1'
                Write-Host "  $formatted" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "  ✗ Error running $($selectedAi['Desc']): $_" -ForegroundColor Red
        Write-Host "    Falling back: report copied to clipboard." -ForegroundColor Yellow
        $aiPrompt | Set-Clipboard
    } finally {
        Remove-Item $tempPrompt -ErrorAction SilentlyContinue
    }

    Write-Host ""
    return
}
