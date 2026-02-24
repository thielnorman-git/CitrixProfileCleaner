<#
.SYNOPSIS
    Citrix Profile Permissions Restore Tool
    
.DESCRIPTION
    Stellt rekursiv Besitz und NTFS-Berechtigungen für Citrix-Profilordner wieder her.
    Optimiert: Prüft vorab, ob Reparatur nötig ist (Owner/ACL Check).
    Inklusive: Long-Path Support, AD-Validierung mit Shift-Logic, Administrator-Check,
    Prüfung auf aktive Dateisperren (User-Login-Check) und Abschluss-Statistik.

.EXAMPLE
    .\RestoreProfilePermissionsTool.ps1 -Path "D:\Profiles"
#>

param (
    # Der Pfad zum Profil oder Root-Ordner (Pflichtfeld)
    [Parameter(Mandatory=$true)]
    [Alias("ProfilePath")]
    [string]$Path,

    # Speicherort für die Log-Dateien (Standard: C:\Logs\ProfileRestore)
    [Parameter(Mandatory=$false)]
    [string]$LogFolder = "C:\Logs\ProfileRestore",

    # Schalter für Automatisierung (unterdrückt Rückfragen und Log-Öffnen)
    [switch]$BatchMode 
)

# --- FUNKTIONEN ---

# Funktion für farbige Konsolenausgaben mit Zeitstempel
function Write-Status {
    param([string]$m, [string]$lvl="INFO")
    
    # Erzeugt Datum/Uhrzeit im Format: 2026-02-24 14:30:00
    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Bestimmt die Farbe basierend auf dem Status-Level
    $color = switch($lvl) {
        "SUCCESS" { "Green" }   # Erfolg
        "ERROR"   { "Red" }     # Fehler
        "WARN"    { "Yellow" }  # Warnung/Suche
        "PROCESS" { "Cyan" }    # Aktive Bearbeitung
        "SKIP"    { "Magenta" } # Gesperrte Profile
        "INFO"    { "Gray" }    # Bereits korrekte Profile
        default   { "White" }
    }
    Write-Host "[$timeStamp] [$lvl] $m" -ForegroundColor $color
}

# Hilfsfunktion: Prüft eine Identität gegen das Active Directory
function Test-ADAccount {
    param([string]$Identity)
    try {
        $ntAccount = New-Object System.Security.Principal.NTAccount($Identity)
        $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
        return $null -ne $sid
    } catch {
        return $false
    }
}

# NEU: Die Shift-Logic zur Identitätsermittlung (inkl. AD-Validierung)
function Get-ValidIdentity {
    param([string]$FolderName)
    
    $parts = $FolderName.Split('.')
    # Wir brauchen mindestens User, Domain, Version
    if ($parts.Count -lt 3) { return $null }

    # Wir ignorieren den letzten Teil (Version, z.B. Win2022v6)
    $meaningfulParts = $parts[0..($parts.Count - 2)]

    # SHIFT-LOGIC: Wir probieren die Trennung an jeder möglichen Punkt-Position
    # Beispiel: becker . j . helios-dom
    for ($i = 1; $i -lt $meaningfulParts.Count; $i++) {
        $potentialUser = ($meaningfulParts[0..($i-1)]) -join "."
        $potentialDomain = $meaningfulParts[$i..($meaningfulParts.Count - 1)] -join "."
        
        $identity = "$potentialDomain\$potentialUser"
        
        # 1. Versuch: Direkte Prüfung (z.B. helios-dom\becker.j oder helios-dom\jbecker)
        if (Test-ADAccount -Identity $identity) {
            return $identity
        }
        
        # 2. Versuch: Flip-Logic (nur bei historischen Schema becker.j -> j.becker)
        if ($potentialUser -contains ".") {
            $subParts = $potentialUser.Split('.')
            if ($subParts.Count -eq 2) {
                $flippedUser = "$($subParts[1]).$($subParts[0])"
                $flippedIdentity = "$potentialDomain\$flippedUser"
                if (Test-ADAccount -Identity $flippedIdentity) {
                    return $flippedIdentity
                }
            }
        }
    }
    return $null
}

# Prüft, ob ein Profil durch eine aktive Nutzersitzung gesperrt ist
function Test-IsFolderLocked {
    param([string]$FolderPath)
    
    # Suche nach der NTUSER.DAT (wird bei Login vom System gesperrt)
    $testFile = Get-ChildItem -Path $FolderPath -Filter "NTUSER.DAT" -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
    
    # Wenn keine NTUSER.DAT da ist, gehen wir davon aus, dass es nicht gesperrt ist
    if ($null -eq $testFile) { return $false } 

    try {
        # Versuche die Datei exklusiv zu öffnen
        $fileStream = [System.IO.File]::Open($testFile.FullName, 'Open', 'Read', 'None')
        $fileStream.Close()
        return $false # Erfolg -> Datei ist frei
    } catch {
        return $true  # Fehler -> Datei ist im Zugriff (User angemeldet)
    }
}

# Prüft vorab, ob Besitzer und Rechte bereits korrekt sind (Smart-Scan)
function Test-NeedsRepair {
    param([string]$FolderPath, [string]$Identity)
    try {
        # Liest die aktuellen Sicherheitsberechtigungen (ACL) aus
        $acl = Get-Acl -Path $FolderPath
        $owner = $acl.Owner # Aktueller Besitzer
        
        # 1. Schritt: Prüfen ob der Besitzername exakt übereinstimmt
        if ($owner -eq $Identity) {
            $hasFullControl = $false
            # 2. Schritt: Prüfen ob die Identität explizit Vollzugriff hat
            foreach($access in $acl.Access) {
                if ($access.IdentityReference -eq $Identity -and $access.FileSystemRights -match "FullControl") {
                    $hasFullControl = $true
                    break
                }
            }
            if ($hasFullControl) { return $false } # Alles korrekt -> keine Reparatur nötig
        }
        return $true # Etwas stimmt nicht -> Reparatur nötig
    } catch {
        return $true # Bei Fehlern (z.B. kein Zugriff auf ACL) Reparatur erzwingen
    }
}

# Kernfunktion: Führt die eigentliche Rechte-Reparatur aus
function Invoke-PermissionReset {
    param([string]$TargetFolder)
    
    $folderName = Split-Path $TargetFolder -Leaf
    
    # NEU: Identität sicher ermitteln (Shift-Logic + AD Check)
    $identity = Get-ValidIdentity -FolderName $folderName
    
    if ($null -eq $identity) {
        Write-Status "AD-FEHLER: '$folderName' konnte nicht aufgelöst werden!" "ERROR"
        return [pscustomobject]@{
            Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            ProfilePath = $TargetFolder
            User        = "UNKNOWN"
            Status      = "Fehler"
            Details     = "Identität nicht im AD gefunden"
        }
    }

    $fullPath = (Get-Item $TargetFolder).FullName
    # Long-Path Support: Präfix \\?\ erlaubt Pfade länger als 260 Zeichen
    $longPath = "\\?\$fullPath"

    # Initialisiert das Ergebnis-Objekt für das Log
    $entry = [pscustomobject]@{
        Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        ProfilePath = $fullPath
        User        = $identity
        Status      = "Geprüft"
        Details     = ""
    }

    # Sicherheit: Überspringen, wenn User noch angemeldet
    if (Test-IsFolderLocked -FolderPath $TargetFolder) {
        Write-Status "ÜBERSPRINGE: $identity (Datei gesperrt)" "SKIP"
        $entry.Status = "Übersprungen"
        $entry.Details = "Datei gesperrt (NTUSER.DAT im Zugriff)"
        return $entry
    }

    # Effizienz: Überspringen, wenn bereits alles okay ist
    if (-not (Test-NeedsRepair -FolderPath $TargetFolder -Identity $identity)) {
        Write-Status "BEREITS OK: $folderName ($identity)" "INFO"
        $entry.Status = "Keine Änderung"
        $entry.Details = "Berechtigungen bereits korrekt"
        return $entry
    }

    # Die eigentliche Reparatur-Kette (External Tools)
    try {
        Write-Status "REPARIERE: $folderName -> $identity" "PROCESS"
        
        # 1. Besitz übernehmen (/f Pfad, /a Administratoren, /r rekursiv, /d y Ja bestätigen)
        & takeown.exe /f "$longPath" /a /r /d y > $null 2>&1
        
        # 2. Besitzer auf den eigentlichen User setzen (/setowner)
        & icacls.exe "$longPath" /setowner "$identity" /t /c /q > $null 2>&1
        
        # 3. Vollzugriff gewähren (OI=ObjectInherit, CI=ContainerInherit, F=Full)
        & icacls.exe "$longPath" /grant "${identity}:(OI)(CI)F" /t /c /q > $null 2>&1
        
        # 4. Vererbung aktivieren (/inheritance:e)
        & icacls.exe "$longPath" /inheritance:e /t /c /q > $null 2>&1

        $entry.Status = "Erfolg"
        Write-Status "ERFOLG: $identity" "SUCCESS"
    } catch {
        # Fehlerbehandlung: Speichert die Fehlermeldung ins Log
        $entry.Status = "Fehler"
        $entry.Details = $_.Exception.Message
        Write-Status "FEHLER bei $folderName" "ERROR"
    }
    return $entry
}

# --- HAUPTPROGRAMM (MAIN) ---

# Prüfung auf Administrator-Rechte beim Start
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Status "Admin-Rechte erforderlich! Bitte PowerShell als Administrator starten." "ERROR"; exit
}

# Prüfen ob der angegebene Pfad existiert
if (-not (Test-Path $Path)) { Write-Status "Pfad nicht gefunden: $Path" "ERROR"; exit }

# Log-Verzeichnis erstellen falls nicht vorhanden
if (-not (Test-Path $LogFolder)) { New-Item $LogFolder -ItemType Directory -Force | Out-Null }
$csvPath = Join-Path $LogFolder "Restore_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Liste zur Speicherung aller Ergebnisse (für das finale CSV-Log)
$report = New-Object System.Collections.Generic.List[PSCustomObject]

# Prüfen ob der Zielpfad ein einzelnes Profil oder ein Root-Ordner ist
$targetDir = Get-Item $Path
$isSingle = $targetDir.Name -match '^.*\..*\..*$'

# Interaktiver Modus: Informationen anzeigen und Bestätigung einholen
if (-not $BatchMode) {
    Write-Host "`n--- CITRIX PROFILE RESTORE PRO v2.3 ---" -ForegroundColor Blue -BackgroundColor White
    Write-Host "Zielpfad:    $Path"
    Write-Host "Log-Datei:   $csvPath"
    Write-Host "Modus:       $(if($isSingle){'Einzelprofil'}else{'Verzeichnis-Scan'})"
    $confirm = Read-Host "`nSollen die Berechtigungen jetzt geprüft/korrigiert werden? (J/N)"
    if ($confirm -ne "J" -and $confirm -ne "j") { Write-Status "Abbruch durch User." "WARN"; exit }
}

# Start der Verarbeitung
if ($isSingle) {
    # Fall 1: Nur ein Ordner angegeben
    $res = Invoke-PermissionReset -TargetFolder $targetDir.FullName
    if ($null -ne $res) { $report.Add($res) }
} else {
    # Fall 2: Root-Verzeichnis scannen
    Write-Status "Scanne Verzeichnis nach Profilen..." "WARN"
    # Findet alle Ordner, die dem Namensschema entsprechen
    $folders = Get-ChildItem -Path $Path -Directory | Where-Object { $_.Name -match '^.*\..*\..*$' }
    
    foreach ($f in $folders) { 
        $report.Add((Invoke-PermissionReset -TargetFolder $f.FullName)) 
    }
}

# --- ABSCHLUSS-STATISTIK ---

# Zählt die verschiedenen Status-Typen in der Ergebnisliste
$total    = $report.Count
$success  = ($report | Where-Object { $_.Status -eq "Erfolg" }).Count
$noChange = ($report | Where-Object { $_.Status -eq "Keine Änderung" }).Count
$skipped  = ($report | Where-Object { $_.Status -eq "Übersprungen" }).Count
$failed   = ($report | Where-Object { $_.Status -eq "Fehler" }).Count

# Ergebnisse als CSV exportieren
$report | Export-Csv -Path $csvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

# Optische Zusammenfassung in der Konsole
Write-Host "`n" + ("=" * 45) -ForegroundColor Cyan
Write-Host "ZUSAMMENFASSUNG DES DURCHLAUFS" -ForegroundColor White
Write-Host ("=" * 45) -ForegroundColor Cyan
Write-Host "Gesamt geprüft:      $total"
Write-Host "Bereits korrekt:     $noChange" -ForegroundColor Gray
Write-Host "Neu repariert:       $success"  -ForegroundColor Green
Write-Host "Übersprungen (Lock): $skipped (Aktiv im Zugriff)" -ForegroundColor Magenta
Write-Host "Fehlerhaft:          $failed"   -ForegroundColor Red
Write-Host ("=" * 45) -ForegroundColor Cyan
Write-Status "Log-Datei erstellt: $csvPath" "SUCCESS"

# Log-Datei automatisch öffnen (nur im interaktiven Modus)
if (-not $BatchMode -and $total -gt 0) {
    $openLog = Read-Host "`nSoll das Logfile (CSV) jetzt geöffnet werden? (J/N)"
    if ($openLog -eq "J" -or $openLog -eq "j") { Invoke-Item $csvPath }
}