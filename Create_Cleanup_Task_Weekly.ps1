<#
.SYNOPSIS
    Create_Cleanup_Task_Weekly.ps1
.DESCRIPTION
    Erstellt eine wöchentliche geplante Aufgabe für den Citrix ProfileCleaner CLI.
    Führt die Bereinigung jeden Sonntag um 02:00 Uhr als SYSTEM aus.
#>

# --- Konfiguration ---
$taskName    = "Citrix_Profile_Cleanup_Weekly"
$scriptPath  = Join-Path $PSScriptRoot "CitrixProfileCleaner_CLI.ps1"
$description = "Wöchentliche automatische Bereinigung von Citrix Benutzerprofilen (2026 Professional Edition)"

# --- 1. Admin-Check ---
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "--------------------------------------------------------" -ForegroundColor Red
    Write-Host "FEHLER: Keine Administratorrechte gefunden!"
    Write-Host "Bitte starten Sie die PowerShell 'Als Administrator', um den Task zu erstellen."
    Write-Host "--------------------------------------------------------"
    exit
}

# --- 2. Prüfung: Existiert das CLI-Skript? ---
if (-not (Test-Path $scriptPath)) {
    Write-Host "FEHLER: Das Skript wurde unter $scriptPath nicht gefunden!" -ForegroundColor Red
    exit
}

# --- 3. Task-Komponenten definieren ---

# Aktion: PowerShell starten und Skript ausführen
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""

# Trigger: Jeden Sonntag um 02:00 Uhr
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2am

# Einstellungen: Aufgabe darf den PC aufwecken, bei Fehlern neu starten
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# --- 4. Task registrieren (mit Fehlerbehandlung) ---
try {
    # Falls der Task schon existiert, wird er ohne Rückfrage überschrieben (-Force)
    Register-ScheduledTask -TaskName $taskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -User "SYSTEM" `
        -RunLevel Highest `
        -Description $description `
        -Force -ErrorAction Stop

    # --- 5. Erfolgsmeldung ---
    Write-Host "--------------------------------------------------------" -ForegroundColor Green
    Write-Host "SUCCESS: Geplante Aufgabe wurde erfolgreich erstellt."
    Write-Host "Taskname:  $taskName"
    Write-Host "Zeitpunkt: Jeden Sonntag um 02:00 Uhr"
    Write-Host "Account:   SYSTEM (Höchste Berechtigungen)"
    Write-Host "Skript:    $scriptPath"
    Write-Host "--------------------------------------------------------" -ForegroundColor Green
}
catch {
    Write-Host "--------------------------------------------------------" -ForegroundColor Red
    Write-Host "FEHLER: Die Aufgabe konnte nicht registriert werden."
    Write-Host "Meldung: $($_.Exception.Message)"
    Write-Host "--------------------------------------------------------"
}