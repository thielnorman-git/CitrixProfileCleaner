<#
.SYNOPSIS
    Citrix Profile Cleaner CLI Runner 2026
.AUTHOR
    Norman Thiel
.VERSION
    1.0 (Meilenstein 30.01.2026)
.DESCRIPTION
    Automatisierter Runner für die ProfileCleanupEngine zur Verwendung in der Konsole.
    Verarbeitet JSON-Jobdefinitionen und steuert den gesamten Prozess inkl. Reporting.
.FEATURES
    - Prüft das 'Enabled'-Flag in den JSON-Jobs zur selektiven Ausführung.
    - Erfasst detaillierte Audit-Infos (takeown/icacls) in der CSV.
    - Zentrales Datei-Logging via Sync-Bridge.
    - Automatischer HTML-Report Trigger nach Abschluss.
.DEPENDENCIES
    - ProfileCleanupEngine.psm1
    - Merge-ProfileCleanerSessionCSVs.psm1
    - Administrator-Rechte
#>

# Ermittlung des Skript-Wurzelverzeichnisses
$projRoot = $PSScriptRoot
if (-not $projRoot) { $projRoot = Get-Location }

# Pfade zu Modulen, Jobs und Logs definieren
$modulePath = Join-Path $projRoot "Modules\ProfileCleanupEngine.psm1"
$reportMod  = Join-Path $projRoot "Modules\Merge-ProfileCleanerSessionCSVs.psm1"
$jobsFolder = Join-Path $projRoot "Jobs"
$lPath      = Join-Path $projRoot "Logs"

# --- SESSION INITIALISIERUNG ---
# Erstellung eines zeitgestempelten Ordners für diese Sitzung
$sessionTimestamp = Get-Date -Format 'yyyyMMdd_HHmm'
$sessionPath = Join-Path $lPath "CLI_Session_$sessionTimestamp"
if (!(Test-Path $sessionPath)) { New-Item $sessionPath -ItemType Directory -Force | Out-Null }

# Globaler Anker für die Log-Datei (wird von der Engine für Write-Log genutzt)
$global:sessionLogFile = Join-Path $sessionPath "Cleanup_Details.log"
"--- CLI SESSION START: $(Get-Date) ---" | Out-File $global:sessionLogFile -Encoding UTF8

# --- CLI Sync-Mockup (Voraussetzung für Engine v2.9) ---
# Erstellt eine kompatible Schnittstelle für die Engine-Logik (Kommunikationsbrücke)
$Sync = @{ 
    Window          = $null; 
    CancelRequested = $false; 
    CurrentJob      = "CLI-Modus";
    ReportData      = New-Object System.Collections.ArrayList;
    LogFile         = $global:sessionLogFile # Ermöglicht der Engine das Schreiben ins Datei-Log
}

# Engine laden und auf Existenz prüfen
if (Test-Path $modulePath) {
    Import-Module $modulePath -Force
} else {
    $err = "[ERROR] Engine Modul nicht gefunden unter: $modulePath"
    Write-Host $err -ForegroundColor Red
    $err | Out-File $global:sessionLogFile -Append
    exit
}

Write-Host "--- CITRIX PROFILE CLEANER CLI START ---" -ForegroundColor Cyan
Write-Host "Log-Datei: $($Sync.LogFile)" -ForegroundColor Gray

# --- Job-Verarbeitung ---
# Alle Job-Definitionen im JSON-Format einlesen und abarbeiten
if (Test-Path $jobsFolder) {
    $jobFiles = Get-ChildItem $jobsFolder -Filter "*.json"
    
    foreach ($file in $jobFiles) {
        try {
            # JSON-Inhalt konvertieren
            $job = Get-Content $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            
            # --- FILTER-LOGIK ---
            # Nur Jobs ausführen, die aktiv geschaltet sind
            if ($job.Enabled -ne $true) {
                $skipMsg = "[SKIP] Job '$($job.Label)' ist deaktiviert (Enabled=false)."
                Write-Host $skipMsg -ForegroundColor Gray
                # Nachricht via Engine-Logik protokollieren
                Write-Log -m $skipMsg -lvl "PASSIV"
                continue
            }

            Write-Host "`n>>> Starte Job: $($job.Label)" -ForegroundColor Yellow
            
            # Aufruf der Kern-Löschfunktion (DryRun:$false für scharfes Löschen im CLI)
            Invoke-ProfileCleanupJob -Job $job -DryRun:$false -Sync $Sync
            
        } catch {
            $errMsg = "[ERROR] Fehler beim Verarbeiten von $($file.Name): $($_.Exception.Message)"
            Write-Host $errMsg -ForegroundColor Red
            if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
                Write-Log -m $errMsg -lvl "ERROR"
            }
        }
    }
}

# --- Reporting & CSV Export ---
# Falls Daten erfasst wurden, Audit-Infos ergänzen und CSV schreiben
if ($Sync.ReportData.Count -gt 0) {
    $auditUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    foreach ($row in $Sync.ReportData) {
        # Setzt den ausführenden User für den Audit-Trail
        if ($null -eq $row.RunBy -or $row.RunBy -eq "") { 
            $row.RunBy = "$auditUser (CLI)"
        }
        # Sicherstellen, dass die Spalte für Audit-Details im Objekt existiert
        if ($null -eq $row.AuditDetails) {
            $row | Add-Member -MemberType NoteProperty -Name "AuditDetails" -Value "" -Force -ErrorAction SilentlyContinue
        }
    }

    # Daten als CSV für den Report-Merger speichern
    $csvPath = Join-Path $sessionPath "Cleanup_Data.csv"
    $Sync.ReportData | Export-Csv $csvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "`n--- REPORTING ---" -ForegroundColor Cyan
    Write-Host "Vollständiger CSV-Log: $csvPath" -ForegroundColor Gray

    # Den grafischen HTML Report über das entsprechende Modul erzeugen
    if (Test-Path $reportMod) {
        Import-Module $reportMod -Force
        Merge-ProfileCleanerSessionCSVs -SessionPath $sessionPath
        Write-Host "HTML-Report wurde erstellt." -ForegroundColor Green
    }
} else {
    Write-Host "`nKeine Daten erfasst. (Möglicherweise keine Jobs aktiv oder keine Profile gefunden)." -ForegroundColor Gray
}

Write-Host "`n--- CLI CLEANUP BEENDET ---" -ForegroundColor Cyan
Write-Host "Alle Details unter: $sessionPath" -ForegroundColor Gray