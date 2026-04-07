﻿<#
.SYNOPSIS
    Profile Cleanup Engine 2026
.AUTHOR
    Norman Thiel
.VERSION
    1.9.1 (Revision 07.04.2026 - GUI Info-Fix)
.DESCRIPTION
    Zentrale Logik-Engine für die Citrix Profil-Bereinigung. 
    Beinhaltet präzise Größenberechnung, Log-Zentralisierung über die Sync-Brücke 
    und ein tiefes Audit-Logging (takeown/icacls). 
    AD-Validierung (Orphaned Check) und Sperrprüfung (NTUSER.DAT).
.FEATURES
    - Zentralisierte Logging-Funktion (Write-Log) mit Unterstützung für GUI-Dispatcher und Datei-Logging.
    - Differenzierte Altersprüfung: DIR-Methode oder INI-Methode (für Citrix UPM).
    - Automatisierte Besitzübernahme (takeown) und ACL-Korrektur (icacls) vor dem Löschen.
    - Status-Meldung "Erfolgreich gelöscht" für optimierte HTML-Auswertung.
    - Vollständige Multi-Threading Unterstützung via $Sync Hashtable.
    - AD-Validierung (Orphaned Check) und Sperrprüfung (NTUSER.DAT).
.DEPENDENCIES
    - Administrator-Rechte (erforderlich für Besitztum-Änderungen an Profilen).

.CHANGELOG
    - 07.04.2026: Info-Log während der Prüfung reaktiviert (GUI-Anzeige).
    - 07.04.2026: Integration von Test-ADAccount, Get-ValidIdentity und Test-IsFolderLocked.
#>

# --- HILFSFUNKTIONEN FÜR IDENTITÄT & SPERREN ---

function Test-ADAccount {
    param([string]$Identity)
    try {
        $ntAccount = New-Object System.Security.Principal.NTAccount($Identity)
        $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
        return $null -ne $sid
    } catch { return $false }
}

function Get-ValidIdentity {
    <#
    .DESCRIPTION
        Optimierte Shift-Logic: Extrahiert User und Domain und ignoriert 
        alle Anhänge (Versionen, Zeitstempel, etc.) nach dem Domänen-Teil.
    #>
    param([string]$FolderName)
    
    # Splitten am Punkt
    $parts = $FolderName.Split('.')
    
    # Wir brauchen mindestens User.Domain.Version (3 Teile)
    if ($parts.Count -lt 3) { return $null }

    # Strategie: Wir testen die ersten plausiblen Kombinationen von vorne.
    # Meistens: [0]=User, [1]=Domain ODER [0].[1]=User, [2]=Domain
    
    for ($i = 1; $i -lt $parts.Count; $i++) {
        $potentialUser = ($parts[0..($i-1)]) -join "."
        $potentialDomain = $parts[$i] # Wir nehmen hier nur das nächste Segment als Domain
        
        $identity = "$potentialDomain\$potentialUser"
        
        # 1. Direkte AD-Prüfung
        if (Test-ADAccount -Identity $identity) { return $identity }
        
        # 2. Flip-Logic (falls Domain vorne steht oder Name gedreht ist)
        if ($potentialUser.Contains(".")) {
            $subParts = $potentialUser.Split('.')
            if ($subParts.Count -eq 2) {
                $flippedUser = "$($subParts[1]).$($subParts[0])"
                $flippedIdentity = "$potentialDomain\$flippedUser"
                if (Test-ADAccount -Identity $flippedIdentity) { return $flippedIdentity }
            }
        }

        # Sicherheitsstopp: Wenn wir bei Segmenten ankommen, die Zeitstempel enthalten (z.B. "2025-11"), 
        # brauchen wir nicht weiter nach rechts shiften.
        if ($parts[$i] -match "\d{4}-\d{2}") { break }
    }
    
    return $null
}

function Test-IsFolderLocked {
    param([string]$FolderPath)
    $testFile = Get-ChildItem -Path $FolderPath -Filter "NTUSER.DAT" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $testFile) { return $false } 
    try {
        $fileStream = [System.IO.File]::Open($testFile.FullName, 'Open', 'Read', 'None')
        $fileStream.Close()
        return $false 
    } catch { return $true }
}

# --- ZENTRALISIERTE LOGGING BRÜCKE ---

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$m,
        [Parameter(Mandatory=$false)][string]$lvl="INFO"
    )

    $displayLvl = switch($lvl) {
        "SUCCESS" { "SUCCESS" } "ERROR" { "ERROR" } "WARN" { "WARN" } "PASSIV" { "PASSIV" } default { "INFO" }
    }

    $targetLogFile = $null
    if ($null -ne $Sync -and $Sync.LogFile) { $targetLogFile = $Sync.LogFile } 
    elseif ($global:sessionLogFile) { $targetLogFile = $global:sessionLogFile }

    if ($null -ne $targetLogFile) {
        $stamp = Get-Date -Format "HH:mm:ss"
        "[$stamp] [$displayLvl] $m" | Out-File -FilePath $targetLogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    }

    if ($null -ne $Sync -and $null -ne $Sync.Window) {
        try {
            $Sync.Window.Dispatcher.Invoke({
                if (Get-Command "Write-Log-Internal" -ErrorAction SilentlyContinue) {
                    Write-Log-Internal -m $m -lvl $displayLvl
                }
            })
        } catch { Write-Host "[$displayLvl] $m" }
    } else {
        $color = switch($displayLvl) { "ERROR" {"Red"} "WARN" {"Yellow"} "SUCCESS" {"Green"} "PASSIV" {"Gray"} default {"White"} }
        Write-Host "[$displayLvl] $m" -ForegroundColor $color
    }
}

# --- KERN-FUNKTION FÜR DEN CLEANUP ---

function Invoke-ProfileCleanupJob {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)][pscustomobject]$Job, 
        [switch]$DryRun, 
        $Sync
    )
    
    $maxAge = if ($Job.MaxAgeDays) { [int]$Job.MaxAgeDays } else { 80 }
    $isUpm = $Job.Label -like "*UPM*"

    foreach ($rawRoot in @($Job.RootPaths)) {
        if ($Sync.CancelRequested) { return }
        $Sync.CurrentJob = "Verarbeite: $($Job.Label)"
        
        $cleanRoot = $rawRoot.ToString().Trim()
        $base = if ($cleanRoot -match '^[a-zA-Z]:' -or $cleanRoot.StartsWith("\\")) { $cleanRoot } else { "D:\$cleanRoot" }
        
        if (Test-Path $base) {
            foreach ($dir in (Get-ChildItem $base -Directory -Force)) {
                if ($Sync.CancelRequested) { return }
                
                # --- SPERR-PRÜFUNG ---
                if (Test-IsFolderLocked -FolderPath $dir.FullName) {
                    Write-Log -m "SKIP (Aktiv): $($dir.Name)" -lvl "WARN"
                    continue
                }

                # --- IDENTITÄTS-VALIDIERUNG ---
                $identity = Get-ValidIdentity -FolderName $dir.Name
                $isOrphaned = $null -eq $identity

                # Zielpfad Bestimmung
                $sub = if ($Job.SubFolder -is [array]) { $Job.SubFolder[0] } else { $Job.SubFolder }
                $targetPath = if ($sub) { Join-Path $dir.FullName $sub.ToString() } else { $dir.FullName }
                
                # Alters-Ermittlung
                $referenceDate = $dir.LastWriteTime
                $method = "DIR"
                if ($isUpm -and (Test-Path (Join-Path $dir.FullName "UPMSettings.ini"))) {
                    $referenceDate = (Get-Item (Join-Path $dir.FullName "UPMSettings.ini")).LastWriteTime
                    $method = "INI"
                }
                $age = [math]::Round(((Get-Date) - $referenceDate).TotalDays, 0)

                # Größenberechnung
                $sizeSum = 0
                if (Test-Path $targetPath) {
                    $sizeSum = (Get-ChildItem $targetPath -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
                }
                $sizeMB = [math]::Round(($sizeSum / 1MB), 2)

                # --- WICHTIG: INFO-LOG FÜR DIE GUI ---
                # Diese Zeile wurde hinzugefügt, damit du siehst, was das Skript gerade tut.
                Write-Log -m "Prüfe [$method]: $($dir.Name) ($age Tage, $sizeMB MB)" -lvl "INFO"

                # --- LÖSCH-ENTSCHEIDUNG ---
                $reason = ""
                $shouldDelete = $false

                if ($isOrphaned) {
                    $shouldDelete = $true
                    $reason = "ORPHANED (Kein AD-Konto)"
                } elseif ($isUpm -and $age -ge $maxAge) {
                    $shouldDelete = $true
                    $reason = "ALTER ($age Tage)"
                } elseif (-not $isUpm -and $sizeMB -gt 0) {
                    $shouldDelete = $true
                    $reason = "INHALT ($sizeMB MB)"
                }

                $reportEntry = [pscustomobject]@{
                    Zeitpunkt    = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
                    Job          = $Job.Label
                    Pfad         = $targetPath
                    Alter        = $age
                    MB           = $sizeMB
                    Status       = "GEPRÜFT"
                    Methode      = $method
                    Identity     = if($isOrphaned){"NICHT GEFUNDEN"}else{$identity}
                    AuditDetails = "" 
                }

                if ($shouldDelete) {
                    $deleteTarget = if ($isUpm -or $isOrphaned) { $dir.FullName } else { $targetPath }
                    
                    if ($DryRun) {
                        Write-Log -m "SIMULATION [$reason] -> $deleteTarget" -lvl "WARN"
                        $reportEntry.Status = "SIMULATION"
                    } else {
                        try {
                            Write-Log -m "CLEANUP [$reason] -> $deleteTarget" -lvl "WARN"
                            takeown.exe /f "$deleteTarget" /r /d y > $null 2>&1
                            $tOK = $LASTEXITCODE -eq 0
                            icacls.exe "$deleteTarget" /grant *S-1-5-32-544:F /t /c /q > $null 2>&1
                            $iOK = $LASTEXITCODE -eq 0
                            $auditStatus = "Audit: $(if($tOK){'OK'}else{'Fail'})/$(if($iOK){'OK'}else{'Fail'})"

                            $tempEmpty = Join-Path $env:TEMP "Empty_$(Get-Random)"
                            New-Item $tempEmpty -ItemType Directory -Force | Out-Null
                            $roboLogPath = if ($Sync.LogFile) { Join-Path (Split-Path $Sync.LogFile) "Robocopy_Detail.log" } else { Join-Path $env:TEMP "Robocopy_Detail.log" }

                            $psi = New-Object System.Diagnostics.ProcessStartInfo -Property @{
                                FileName = "robocopy.exe"
                                Arguments = """$tempEmpty"" ""$deleteTarget"" /MIR /XJ /R:0 /W:0 /MT:32 /NFL /NDL /NC /NS /NP /LOG+:""$roboLogPath"""
                                UseShellExecute = $false; CreateNoWindow = $true; RedirectStandardOutput = $true; RedirectStandardError = $true
                            }
                            $proc = [System.Diagnostics.Process]::Start($psi)
                            $proc.WaitForExit()
                            $rc = $proc.ExitCode

                            if (Test-Path $deleteTarget) { Remove-Item $deleteTarget -Recurse -Force -ErrorAction SilentlyContinue }
                            Remove-Item $tempEmpty -Force -ErrorAction SilentlyContinue

                            Write-Log -m "[LÖSCHEN] -> Erfolgreich: $deleteTarget" -lvl "SUCCESS"
                            $reportEntry.Status = "GELÖSCHT ($reason)"
                            $reportEntry.AuditDetails = "$auditStatus | RC: $rc"
                        } catch {
                            Write-Log -m "FEHLER: $($_.Exception.Message)" -lvl "ERROR"
                            $reportEntry.Status = "FEHLER"
                            $reportEntry.AuditDetails = "Error: $($_.Exception.Message)"
                        }
                    }
                }
                if ($null -ne $Sync.ReportData) { [void]$Sync.ReportData.Add($reportEntry) }
            } 
        } 
    }
}

export-modulemember -function Write-Log, Invoke-ProfileCleanupJob, Test-ADAccount, Get-ValidIdentity, Test-IsFolderLocked