<#
.SYNOPSIS
Collects AD computer info and (optionally) detailed data in parallel:
 - OS, AD OU, AD update time
 - Ping status
 - Optional: Uptime, LastUser, and LastPatch info
Uses PowerShell background jobs with 60-second timeout per host.
Marks jobs that exceed timeout.
#>

#Requires -Modules ActiveDirectory

# --- SETTINGS ---

If ($psISE) {

    $CurrentDir = Split-path $psISE.CurrentFile.FullPath
}

Else {

    $CurrentDir = split-path -parent $MyInvocation.MyCommand.Definition

}

If (-not(test-path "$CurrentDir\Reports")) {

    New-item -Path "$CurrentDir\Reports" -ItemType Directory

}

$XLSExportPath = "$CurrentDir\Get-ADComputers-LogTimeStamp.xlsx"
$JobTimeoutSeconds = 60
$MaxConcurrentJobs = 20

# --- MODE SELECTION ---
Write-Host "===============================" -ForegroundColor Cyan
Write-Host " AD Computer Audit Script" -ForegroundColor White
Write-Host "===============================" -ForegroundColor Cyan
Write-Host "`nThis script can collect:" -ForegroundColor Gray
Write-Host "  • Computer name, OS, AD OU, AD update time" -ForegroundColor DarkGray
Write-Host "  • Ping status" -ForegroundColor DarkGray
Write-Host "  • (Optional) Last user, uptime, and last patch info" -ForegroundColor DarkGray
Write-Host "`nIf you choose 'Ping Only' mode, these details will be SKIPPED:" -ForegroundColor Yellow
Write-Host "  ✗ Last user info" -ForegroundColor Yellow
Write-Host "  ✗ Uptime" -ForegroundColor Yellow
Write-Host "  ✗ Patch information" -ForegroundColor Yellow
Write-Host "→ Runs significantly faster." -ForegroundColor Gray

$pingOnlyChoice = Read-Host "`nDo you want to run in PING ONLY mode? (Y/N)"
$PingOnly = $pingOnlyChoice -match '^[Yy]'

if ($PingOnly) {
    Write-Host "`n⚡ Running in PING ONLY mode — skipping WMI data collection." -ForegroundColor Yellow
} else {
    Write-Host "`n🧩 Full mode selected — collecting detailed info via WMI." -ForegroundColor Green
}

# --- STEP 1: Get all AD computer objects ---
Write-Host "`n[1/5] Collecting AD computer objects from domain..." -ForegroundColor Cyan
$AllADComputers = Get-ADComputer -Filter * -Property Name,OperatingSystem,whenChanged,DistinguishedName |
    Select-Object Name,
                  @{Name='OperatingSystem';Expression={$_.OperatingSystem}},
                  @{Name='LastObjectUpdate';Expression={$_.whenChanged}},
                  @{Name='OU';Expression={($_.DistinguishedName -split ',(?=OU=)') -join '/'}}

Write-Host "Collected $($AllADComputers.Count) computers from AD." -ForegroundColor Green

# --- STEP 2: Test connectivity ---
Write-Host "`n[2/5] Testing connectivity for Windows computers..." -ForegroundColor Cyan
$WindowsComputers = $AllADComputers | Where-Object { $_.OperatingSystem -match 'Windows' }

$total = $WindowsComputers.Count
$counter = 0

foreach ($comp in $WindowsComputers) {
    $counter++
    Write-Progress -Activity "Pinging Windows Computers" -Status "$counter / $total" -PercentComplete (($counter / $total) * 100)
    Write-Host ("Pinging {0} ({1}/{2})..." -f $comp.Name, $counter, $total) -ForegroundColor DarkCyan

    $isOnline = Test-Connection -ComputerName $comp.Name -Count 1 -Quiet -ErrorAction SilentlyContinue
    $comp | Add-Member -NotePropertyName Online -NotePropertyValue $isOnline -Force
    if ($isOnline) {
        Write-Host ("  ✓ {0} is online" -f $comp.Name) -ForegroundColor Green
    } else {
        Write-Host ("  ✗ {0} is offline or unreachable" -f $comp.Name) -ForegroundColor DarkGray
    }
}

$OnlineWindows = $WindowsComputers | Where-Object { $_.Online -eq $true }
Write-Host "`nOnline Windows computers: $($OnlineWindows.Count)" -ForegroundColor Green

# --- STEP 3: Detailed WMI info via jobs ---
$Results = @()

if (-not $PingOnly -and $OnlineWindows.Count -gt 0) {
    Write-Host "`n[3/5] Gathering detailed info in parallel (max $MaxConcurrentJobs jobs)..." -ForegroundColor Cyan
    $JobList = @()

    foreach ($comp in $OnlineWindows) {
        while (@(Get-Job -State Running).Count -ge $MaxConcurrentJobs) {
            Start-Sleep -Seconds 1
        }

        Write-Host ("Starting job for {0}" -f $comp.Name) -ForegroundColor Cyan
        $job = Start-Job -Name $comp.Name -ArgumentList $comp.Name -ScriptBlock {
            param($ComputerName)

            $result = [ordered]@{
                Name              = $ComputerName
                LastUser          = "Unknown"
                LastUserLogonTime = $null
                UptimeDays        = $null
                LastPatchID       = $null
                LastPatchDate     = $null
                Error             = $null
                TimedOut          = $false
            }

            try {
                $CimSession = New-CimSession -ComputerName $ComputerName -ErrorAction Stop

                # --- Last user ---
                try {
                    $scriptBlock = {
                        if (Test-Path 'C:\Users') {
                            Get-ChildItem 'C:\Users' -Directory |
                            Where-Object { $_.Name -notmatch 'Public|Default' } |
                            Sort-Object LastWriteTime -Descending |
                            Select-Object -First 1 -Property Name, LastWriteTime
                        }
                    }
                    $lastUserData = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ErrorAction SilentlyContinue
                    if ($lastUserData) {
                        $result.LastUser = $lastUserData.Name
                        $result.LastUserLogonTime = $lastUserData.LastWriteTime
                    }
                } catch {}

                # --- Uptime ---
                try {
                    $os = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $CimSession -ErrorAction Stop
                    $uptime = (Get-Date) - $os.LastBootUpTime
                    $result.UptimeDays = [math]::Round($uptime.TotalDays,2)
                } catch {}

                # --- Patch info ---
                try {
                    $lastPatch = Get-CimInstance -ClassName Win32_QuickFixEngineering -CimSession $CimSession -ErrorAction SilentlyContinue |
                                 Sort-Object -Property InstalledOn -Descending |
                                 Select-Object -First 1 -Property HotFixID,InstalledOn
                    if ($lastPatch) {
                        $result.LastPatchID = $lastPatch.HotFixID
                        $result.LastPatchDate = $lastPatch.InstalledOn
                    }
                } catch {}

                Remove-CimSession $CimSession
            } catch {
                $result.Error = $_.Exception.Message
            }

            return [pscustomobject]$result
        }
        $JobList += $job
    }

    # --- Monitor jobs & handle timeouts ---
    Write-Host "`nWaiting for all jobs to complete (timeout $JobTimeoutSeconds sec each)..." -ForegroundColor Gray
    while (@(Get-Job -State Running).Count -gt 0) {
        $running = @(Get-Job -State Running)
        Write-Progress -Activity "Collecting WMI info" -Status "$($running.Count) jobs running..." -PercentComplete ((($OnlineWindows.Count - $running.Count) / $OnlineWindows.Count) * 100)
        Start-Sleep -Seconds 2

        foreach ($job in $running) {
            $elapsed = (Get-Date) - $job.PSBeginTime
            if ($elapsed.TotalSeconds -ge $JobTimeoutSeconds) {
                Write-Warning "Timeout reached for $($job.Name) — marking as TimedOut."
                Stop-Job $job | Out-Null

                # Create a placeholder result
                $timeoutResult = [pscustomobject]@{
                    Name              = $job.Name
                    LastUser          = $null
                    LastUserLogonTime = $null
                    UptimeDays        = $null
                    LastPatchID       = $null
                    LastPatchDate     = $null
                    Error             = "Job exceeded $JobTimeoutSeconds sec timeout."
                    TimedOut          = $true
                }
                $Results += $timeoutResult

                Receive-Job $job -ErrorAction SilentlyContinue | Out-Null
                Remove-Job $job
            }
        }
    }

    # --- Gather finished job results ---
    Write-Host "`nCollecting job results..." -ForegroundColor Cyan
    foreach ($job in Get-Job) {
        try {
            $data = Receive-Job $job -ErrorAction SilentlyContinue
            if ($data) { $Results += $data }
        } catch {}
        Remove-Job $job
    }

    # --- Merge results ---
    foreach ($comp in $OnlineWindows) {
        $r = $Results | Where-Object { $_.Name -eq $comp.Name }
        if ($r) {
            foreach ($prop in $r.PSObject.Properties) {
                $comp | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
            }
        }
    }
} else {
    Write-Host "`n[3/5] Skipping WMI data collection (ping-only or no online systems)." -ForegroundColor Yellow
}

# --- STEP 4: Export results ---
Write-Host "`n[4/5] Exporting results..." -ForegroundColor Cyan
if (Get-Module -ListAvailable -Name ImportExcel) {
    $AllADComputers | Export-Excel -Path $XLSReport -AutoSize -FreezeTopRow
    Write-Host "✅ Results exported to Excel: $XLSReport" -ForegroundColor Green
} else {
    $CsvPath = $XLSReport -replace '\.xlsx$', '.csv'
    $AllADComputers | Export-Csv -Path $CsvPath -NoTypeInformation
    Write-Host "⚠️ ImportExcel module not found. Exported to CSV instead: $CsvPath" -ForegroundColor Yellow
}

# --- STEP 5: Out-GridView ---
Write-Host "`n[5/5] Opening results in Out-GridView..." -ForegroundColor Cyan
$AllADComputers | Out-GridView -Title "AD Computer Audit Results for $env:USERDNSDOMAIN"

Write-Host "`n✅ Script completed successfully." -ForegroundColor Green
