<#
.SYNOPSIS
    Profile Cleanup Engine 2026
.AUTHOR
    Norman Thiel
.VERSION
    1.7 (Revision 20.02.2026)
.DESCRIPTION
    Zentrale Logik-Engine für die Citrix Profil-Bereinigung. 
    Beinhaltet präzise Größenberechnung, Log-Zentralisierung über die Sync-Brücke 
    und ein tiefes Audit-Logging (takeown/icacls).
.CHANGELOG
    - 20.02.2026: Löschprozess umgestellt von .NET auf Robocopy.
    - Audit- und Robocopy-Statusmeldungen in Logausgabe hinzugefügt.
    - TakeOwn und ACL Übernahme werden jetzt nur noch auf den Ordner der auch wirklich gelöscht wird ausgeführt.
.FEATURES
    - Zentralisierte Logging-Funktion (Write-Log) mit Unterstützung für GUI-Dispatcher.
    - Automatisierte Besitzübernahme (takeown) und ACL-Korrektur (icacls).
    - Vollständige Multi-Threading Unterstützung via $Sync Hashtable.
.DEPENDENCIES
    - Administrator-Rechte (erforderlich für Besitztum-Änderungen an Profilen).
#>

# --- ZENTRALISIERTE LOGGING BRÜCKE ---
function Write-Log {
    <#
    .DESCRIPTION
        Schreibt Log-Einträge parallel in:
        1. Die GUI (via Dispatcher-Invoke, falls vorhanden)
        2. Die Konsolenausgabe (Write-Host mit Farbcodierung)
        3. Die zentrale Log-Datei der aktuellen Session.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$m,
        [Parameter(Mandatory=$false)][string]$lvl="INFO"
    )

    # Vereinheitlichung der Log-Level für Datei- und Konsolenausgabe
    $displayLvl = switch($lvl) {
        "SUCCESS" { "SUCCESS" } 
        "ERROR"   { "ERROR" }
        "WARN"    { "WARN" }
        "PASSIV"  { "PASSIV" }
        default   { "INFO" }
    }

    # Ermittlung des Log-Pfads (Entweder über Sync-Objekt aus dem Thread oder globalen Anker)
    $targetLogFile = $null
    if ($null -ne $Sync -and $Sync.LogFile) { 
        $targetLogFile = $Sync.LogFile 
    } elseif ($global:sessionLogFile) { 
        $targetLogFile = $global:sessionLogFile 
    }

    # Schreiben in die Datei (UTF8 für Sonderzeichen in Pfaden)
    if ($null -ne $targetLogFile) {
        $stamp = Get-Date -Format "HH:mm:ss"
        "[$stamp] [$displayLvl] $m" | Out-File -FilePath $targetLogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    }

    # Fallunterscheidung: GUI-Update oder Konsolenausgabe
    if ($null -ne $Sync -and $null -ne $Sync.Window) {
        try {
            # Sicherer Thread-Zugriff auf das WPF-Fenster (Thread-Safe Dispatcher)
            $Sync.Window.Dispatcher.Invoke({
                if (Get-Command "Write-Log-Internal" -ErrorAction SilentlyContinue) {
                    Write-Log-Internal -m $m -lvl $displayLvl
                }
            })
        } catch {
            Write-Host "[$displayLvl] $m"
        }
    } 
    else {
        # Standard Konsolenausgabe mit Farben für Standalone-Betrieb
        $color = switch($displayLvl) {
            "ERROR"   { "Red" }
            "WARN"    { "Yellow" }
            "SUCCESS" { "Green" }
            "PASSIV"  { "Gray" }
            default   { "White" }
        }
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
    
    # Standard-Alter festlegen (Fallback auf 80 Tage)
    $maxAge = if ($Job.MaxAgeDays) { [int]$Job.MaxAgeDays } else { 80 }
    $isUpm = $Job.Label -like "*UPM*"

    # Verarbeitung aller definierten Pfad-Wurzeln (RootPaths) aus dem JSON
    foreach ($rawRoot in @($Job.RootPaths)) {
        if ($Sync.CancelRequested) { return }
        $Sync.CurrentJob = "Verarbeite: $($Job.Label)"
        
        # Pfad-Validierung und automatische Ergänzung bei relativen Pfaden (D:\ Standard)
        $cleanRoot = $rawRoot.ToString().Trim()
        $base = if ($cleanRoot -match '^[a-zA-Z]:' -or $cleanRoot.StartsWith("\\")) { $cleanRoot } else { "D:\$cleanRoot" }
        
        if (Test-Path $base) {
            # Iteration durch alle Profilordner (Directory Only)
            foreach ($dir in (Get-ChildItem $base -Directory -Force)) {
                if ($Sync.CancelRequested) { return }
                
                # Zielpfad innerhalb des Profils bestimmen (SubFolder aus Konfig)
                $sub = if ($Job.SubFolder -is [array]) { $Job.SubFolder[0] } else { $Job.SubFolder }
                $targetPath = if ($sub) { Join-Path $dir.FullName $sub.ToString() } else { $dir.FullName }
                
                # --- ALTERS-ERMITTLUNG (INI vs DIR) ---
                $referenceDate = $dir.LastWriteTime
                $method = "DIR"
                if ($isUpm -and (Test-Path (Join-Path $dir.FullName "UPMSettings.ini"))) {
                    # Präzise Methode für Citrix UPM: Letzter Schreibzugriff auf die INI-Zentraldatei
                    $referenceDate = (Get-Item (Join-Path $dir.FullName "UPMSettings.ini")).LastWriteTime
                    $method = "INI"
                }

                $age = [math]::Round(((Get-Date) - $referenceDate).TotalDays, 0)

                if (Test-Path $targetPath) {
                    # Lösch-Target festlegen (UPM löscht gesamtes Profil, sonst nur Subfolder-Inhalt)
                    $deleteTarget = if ($isUpm) { $dir.FullName } else { $targetPath }
                    $calcPath = if ($isUpm) { $dir.FullName } else { $targetPath }
                    
                    # --- GRÖSSENBERECHNUNG ---
                    $sizeSum = (Get-ChildItem $calcPath -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
                    $sizeMB = [math]::Round(($sizeSum / 1MB), 2)
                    
                    Write-Log -m "Prüfe [$method]: $targetPath ($age Tage, $sizeMB MB)" -lvl "PASSIV"

                    # Report-Datenobjekt für CSV/HTML initialisieren
                    $reportEntry = [pscustomobject]@{
                        Zeitpunkt    = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
                        Job          = $Job.Label
                        Pfad         = $targetPath
                        Alter        = $age
                        MB           = $sizeMB
                        Status       = "GEPRÜFT"
                        Methode      = $method
                        AuditDetails = "" 
                    }

                    # Löschlogik: UPM via Alter, SubFolder via Existenz von Daten (MB > 0)
                    $shouldDelete = if ($isUpm) { $age -ge $maxAge } else { $sizeMB -gt 0 }

                    if ($shouldDelete) {
                        if ($DryRun) {
                            Write-Log -m "SIMULATION -> Würde bereinigen: $deleteTarget" -lvl "WARN"
                            $reportEntry.Status = "SIMULATION"
                        } else {
                            try {
                                # --- 1. ZIELGERICHTETES AUDIT ---
                                $ownerTarget = $deleteTarget
                                Write-Log -m "Audit: Takeown & ACLs für $ownerTarget" -lvl "WARN"
                                
                                # Besitz übernehmen (Takeown)
                                takeown.exe /f "$ownerTarget" /r /d y > $null 2>&1
                                $tOK = $LASTEXITCODE -eq 0
                                
                                # Administratoren Vollzugriff gewähren (Icacls)
                                icacls.exe "$ownerTarget" /grant *S-1-5-32-544:F /t /c /q > $null 2>&1
                                $iOK = $LASTEXITCODE -eq 0
                                
                                $auditStatus = "Audit: $(if($tOK){'OK'}else{'Fail'})/$(if($iOK){'OK'}else{'Fail'})"

                                # --- 2. LÖSCHVORGANG VIA ROBOCOPY ---
                                # Leerer Quellordner für die Spiegelung erzeugen
                                $tempEmpty = Join-Path $env:TEMP "Empty_$(Get-Random)"
                                New-Item $tempEmpty -ItemType Directory -Force | Out-Null

                                # Log-Pfad aus Sync-Objekt oder Temp
                                $roboLogPath = if ($Sync.LogFile) { Join-Path (Split-Path $Sync.LogFile) "Robocopy_Detail.log" } else { Join-Path $env:TEMP "Robocopy_Detail.log" }

                                # Argumente für die Log-Datei
                                $roboArgs = @("""$tempEmpty""", """$ownerTarget""", "/MIR", "/XJ", "/R:0", "/W:0", "/MT:32", "/NFL", "/NDL", "/NC", "/NS", "/NP", "/LOG+:""$roboLogPath""")
                                
                                Write-Log -m "Robocopy-Löschung läuft.." -lvl "WARN"
                                
                                # VOLLSTÄNDIGE STUMMSCHALTUNG: .NET Process Object fängt den Stream ab
                                $psi = New-Object System.Diagnostics.ProcessStartInfo
                                $psi.FileName = "robocopy.exe"
                                $psi.Arguments = $roboArgs -join " "
                                $psi.UseShellExecute = $false
                                $psi.CreateNoWindow = $true
                                $psi.RedirectStandardOutput = $true # Schickt Output ins Nirvana statt in die Konsole
                                $psi.RedirectStandardError = $true
                                
                                $proc = [System.Diagnostics.Process]::Start($psi)
                                $proc.WaitForExit()

                                # ExitCode-Übersetzung für den Report
                                $rc = $proc.ExitCode
                                $rcDesc = switch($rc) {
                                    0 { "OK" }
                                    1 { "OK" }
                                    2 { "Gelöscht" }
                                    4 { "Geändert" }
                                    8 { "Teilfehler" }
                                    16 { "Pfadfehler" }
                                    default { "Code $rc" }
                                }

                                Write-Log -m "Robocopy beendet: $rcDesc ($rc)" -lvl "WARN"

                                # Finale Aufräumarbeiten (Reste entfernen und Temp-Ordner löschen)
                                if (Test-Path $ownerTarget) {
                                    Remove-Item $ownerTarget -Recurse -Force -ErrorAction SilentlyContinue
                                }
                                Remove-Item $tempEmpty -Force -ErrorAction SilentlyContinue

                                Write-Log -m "[LÖSCHEN] -> Bereinigt: $ownerTarget" -lvl "SUCCESS"
                                
                                # Finale Status-Zuweisung für den HTML-Report
                                $reportEntry.Status = "Erfolgreich gelöscht"
                                $reportEntry.AuditDetails = "$auditStatus | Robocopy: $rcDesc ($rc)"
                            } catch {
                                Write-Log -m "FEHLER: $targetPath - $($_.Exception.Message)" -lvl "ERROR"
                                $reportEntry.Status = "FEHLER"
                                $reportEntry.AuditDetails = "Error: $($_.Exception.Message)"
                            }
                        }
                    }
                    # Datenpunkt zur Thread-sicheren Liste für den Report-Generator hinzufügen
                    if ($null -ne $Sync.ReportData) { [void]$Sync.ReportData.Add($reportEntry) }
                } 
            } 
        } 
    }
}

# Export der Funktionen für das Hauptskript
export-modulemember -function Write-Log, Invoke-ProfileCleanupJob