<#
.SYNOPSIS
    Profile Cleanup Engine 2026
.AUTHOR
    Norman Thiel
.VERSION
    1.0 (Meilenstein 30.01.2026)
.DESCRIPTION
    Zentrale Logik-Engine für die Citrix Profil-Bereinigung. 
    Beinhaltet präzise Größenberechnung, Log-Zentralisierung über die Sync-Brücke 
    und ein tiefes Audit-Logging (takeown/icacls).
.FEATURES
    - Zentralisierte Logging-Funktion (Write-Log) mit Unterstützung für GUI-Dispatcher und Datei-Logging.
    - Differenzierte Altersprüfung: DIR-Methode oder INI-Methode (für Citrix UPM).
    - Automatisierte Besitzübernahme (takeown) und ACL-Korrektur (icacls) vor dem Löschen.
    - Status-Meldung "Erfolgreich gelöscht" für optimierte HTML-Auswertung.
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

    # Vereinheitlichung der Log-Level
    $displayLvl = switch($lvl) {
        "SUCCESS" { "SUCCESS" } 
        "ERROR"   { "ERROR" }
        "WARN"    { "WARN" }
        "PASSIV"  { "PASSIV" }
        default   { "INFO" }
    }

    # Ermittlung des Log-Pfads (Entweder über Sync-Objekt oder globalen Anker)
    $targetLogFile = $null
    if ($null -ne $Sync -and $Sync.LogFile) { 
        $targetLogFile = $Sync.LogFile 
    } elseif ($global:sessionLogFile) { 
        $targetLogFile = $global:sessionLogFile 
    }

    # Schreiben in die Datei
    if ($null -ne $targetLogFile) {
        $stamp = Get-Date -Format "HH:mm:ss"
        "[$stamp] [$displayLvl] $m" | Out-File -FilePath $targetLogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    }

    # Fallunterscheidung: GUI-Update oder Konsolenausgabe
    if ($null -ne $Sync -and $null -ne $Sync.Window) {
        try {
            # Sicherer Thread-Zugriff auf das WPF-Fenster
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
        # Standard Konsolenausgabe mit Farben
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
    
    # Standard-Alter festlegen, falls im Job-JSON nicht definiert
    $maxAge = if ($Job.MaxAgeDays) { [int]$Job.MaxAgeDays } else { 80 }
    $isUpm = $Job.Label -like "*UPM*"

    # Verarbeitung aller definierten Pfad-Wurzeln (RootPaths)
    foreach ($rawRoot in @($Job.RootPaths)) {
        if ($Sync.CancelRequested) { return }
        $Sync.CurrentJob = "Verarbeite: $($Job.Label)"
        
        # Pfad-Validierung und automatische Ergänzung bei relativen Pfaden
        $cleanRoot = $rawRoot.ToString().Trim()
        $base = if ($cleanRoot -match '^[a-zA-Z]:' -or $cleanRoot.StartsWith("\\")) { $cleanRoot } else { "D:\$cleanRoot" }
        
        if (Test-Path $base) {
            # Suche nach Profilordnern in der Basis
            foreach ($dir in (Get-ChildItem $base -Directory -Force)) {
                if ($Sync.CancelRequested) { return }
                
                # Zielpfad innerhalb des Profils bestimmen (SubFolder)
                $sub = if ($Job.SubFolder -is [array]) { $Job.SubFolder[0] } else { $Job.SubFolder }
                $targetPath = if ($sub) { Join-Path $dir.FullName $sub.ToString() } else { $dir.FullName }
                
                # --- ALTERS-ERMITTLUNG (INI vs DIR) ---
                $referenceDate = $dir.LastWriteTime
                $method = "DIR"
                if ($isUpm -and (Test-Path (Join-Path $dir.FullName "UPMSettings.ini"))) {
                    # Präzise Methode für Citrix UPM: Letzter Schreibzugriff auf die INI-Datei
                    $referenceDate = (Get-Item (Join-Path $dir.FullName "UPMSettings.ini")).LastWriteTime
                    $method = "INI"
                }

                $age = [math]::Round(((Get-Date) - $referenceDate).TotalDays, 0)

                if (Test-Path $targetPath) {
                    # Lösch-Target festlegen (UPM löscht den ganzen Ordner, sonst Inhalt des Subfolders)
                    $deleteTarget = if ($isUpm) { $dir.FullName } else { Join-Path $targetPath "*" }
                    $calcPath = if ($isUpm) { $dir.FullName } else { $targetPath }
                    
                    # --- GRÖSSENBERECHNUNG  ---
                    $sizeSum = (Get-ChildItem $calcPath -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
                    $sizeMB = [math]::Round(($sizeSum / 1MB), 2)
                    
                    Write-Log -m "Prüfe [$method]: $targetPath ($age Tage, $sizeMB MB)" -lvl "PASSIV"

                    # Report-Datenobjekt initialisieren
                    $reportEntry = [pscustomobject]@{
                        Zeitpunkt    = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
                        Job          = $Job.Label
                        Pfad         = $targetPath
                        Alter        = $age
                        MB           = $sizeMB
                        Status       = "GEPRÜFT"
                        Methode      = $method
                        RunBy        = ""  
                        AuditDetails = "" 
                    }

                    # Entscheidung: Soll gelöscht werden? (UPM prüft Alter, SubFolder-Jobs löschen immer)
                    $shouldDelete = if ($isUpm) { $age -ge $maxAge } else { $true }

                    if ($shouldDelete) {
                        if ($DryRun) {
                            Write-Log -m "SIMULATION -> Würde bereinigen: $deleteTarget" -lvl "WARN"
                            $reportEntry.Status = "SIMULATION"
                        } else {
                            try {
                                # --- BESITZÜBERNAHME & AUDIT ---
                                # Wir übernehmen den Besitz des gesamten Profilordners für volle Kontrolle
                                $ownerTarget = $dir.FullName
                                
                                $resTakeown = (takeown.exe /f "$ownerTarget" /r /d y 2>&1 | Out-String).Trim()
                                Write-Log -m "Audit (Takeown): $resTakeown" -lvl "PASSIV"
                                
                                # Setzen der NTFS-Berechtigungen für Administratoren
                                $resIcacls = (icacls.exe "$ownerTarget" /grant *S-1-5-32-544:F /t /c /q 2>&1 | Out-String).Trim()
                                Write-Log -m "Audit (ACLs): $resIcacls" -lvl "PASSIV"
                                
                                $auditOutput = "OWNER: $resTakeown | ACLs: $resIcacls"

                                # --- LÖSCHVORGANG ---
                                Remove-Item $deleteTarget -Recurse -Force -ErrorAction Stop
                                
                                Write-Log -m "[LÖSCHEN] -> Bereinigt: $deleteTarget" -lvl "SUCCESS"
                                
                                # Status-Text für die HTML-Berichterstattung setzen
                                $reportEntry.Status = "Erfolgreich gelöscht"
                                $reportEntry.AuditDetails = $auditOutput
                            } catch {
                                Write-Log -m "FEHLER: $targetPath - $($_.Exception.Message)" -lvl "ERROR"
                                $reportEntry.Status = "FEHLER"
                                $reportEntry.AuditDetails = "System-Audit: $auditOutput | Error: $($_.Exception.Message)"
                            }
                        }
                    }
                    # Datenpunkt zur Session-Liste hinzufügen
                    if ($null -ne $Sync.ReportData) { [void]$Sync.ReportData.Add($reportEntry) }
                } 
            } 
        } 
    }
}

# Modul-Schnittstellen exportieren
export-modulemember -function Write-Log, Invoke-ProfileCleanupJob