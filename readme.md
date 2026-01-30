Profile Cleaner 2026 (Gesamtdokumentation)

Ein hybrides Administrations-Tool zur effizienten Verwaltung und Bereinigung von Benutzerprofilen und Verzeichnissen. Das Projekt bietet sowohl eine interaktive WPF-Oberfläche für manuelle Eingriffe als auch eine CLI-Schnittstelle für automatisierte Abläufe.

📂 Projektstruktur
Die Struktur wurde für maximale Portabilität optimiert. Alle Pfade werden relativ zum Skriptverzeichnis aufgelöst.

```text
CitrixProfileCleaner/
├── CitrixProfileCleaner_GUI.ps1     # Haupteinstiegspunkt (WPF-Oberfläche)
├── CitrixProfileCleaner_CLI.ps1     # Autarker Entrypoint für Scheduled Jobs (CLI)
│
├── Modules/
│   ├── ProfileCleanupEngine.psm1    # Kern-Logik: Löschprozesse & Altersprüfung
│   └── Merge-ProfileCleanerSessionCSVs.psm1 # Report-Generator (LOGS/HTML)
│
├── Jobs/                            # JSON-Aufgabenbeschreibungen (Vollpfade)
└── Logs/                            # Sitzungsprotokolle (CSV) & Berichte (HTML)
```

⚙️ Funktionsweise der Engine
Die Engine verarbeitet Vollpfade (RootPaths), die direkt in den Job-Dateien definiert sind. Ein manuelles Auswählen eines Basisverzeichnisses ist nicht erforderlich.

1. Citrix UPM Profile (Type: "UPMCleanup")
Ziel: Vollständige Entfernung alter Profilverzeichnisse zur Speicherplatzrückgewinnung.

Prüfung: Primär wird die UPMSettings.ini im Profil ausgelesen.

Aktion: Wenn das Alter >= MaxAgeDays ist, wird das gesamte Profilverzeichnis gelöscht.

Sicherheit: Inkludiert automatische Rechteübernahme für blockierte Profile.

🛠 Konfiguration (JSON-Jobs)
Die Jobs definieren ihre Ziele über absolute Pfade.

```text
Parameter    Typ        Beschreibung
Label        String     Anzeigename der Aufgabe in der GUI.
Type         String     UPMCleanup (Profil-Logik) oder ProfileFolder (Inhalt löschen).
RootPaths    Array      Vollständige Pfade zu den Profil-Speichern.
SubFolder    String     Relativer Pfad zum Zielordner (nur bei ProfileFolder).
MaxAgeDays   Integer    Schwellenwert für die Löschung in Tagen.
Enabled      Boolean    Schaltet den Job aktiv (true) oder inaktiv (false).
```

JSON ConfigFiles

Beispiel: Template_UPMCleanup_Profile.json
```text
JSON
{
    "Label": "VORLAGE: Citrix UPM Profile (30 Tage)",
    "Type": "UPMCleanup",
    "RootPaths": [
        "\\\\Server01\\CtxProfiles$",
        "\\\\Server02\\CtxProfiles$"
    ],
    "MaxAgeDays": 30,
    "Enabled": true,
    "Comment": "Löscht das gesamte Profilverzeichnis, wenn der Logout länger als 30 Tage her ist."
}
```

Beispiel: Template_Folder.json
```text
JSON
{
    "Label": "VORLAGE: Teams Cache Bereinigung",
    "Type": "ProfileFolder",
    "RootPaths": [
        "\\\\Server01\\CtxProfiles$"
    ],
    "SubFolder": "AppData\\Roaming\\Microsoft\\Teams\\Cache",
    "MaxAgeDays": 0,
    "Enabled": false,
    "Comment": "Löscht nur den Inhalt des SubFolders."
}
```

🚀 Nutzung & Automatisierung

Manueller Modus (GUI)

Start: Rechtsklick auf CitrixProfileCleaner_GUI.ps1 -> Mit PowerShell als Administrator ausführen.

Features: Live-Log-Filter (INFO, WARN, ERROR), Simulationsmodus (Dry-Run) standardmäßig aktiv.

Automatisierter Modus (Scheduled Task)

Skript: CitrixProfileCleaner_CLI.ps1


Task-Konfiguration:

Programm/Skript: powershell.exe

Argumente: -NoProfile -ExecutionPolicy Bypass -File "C:\Pfad\Zu\CitrixProfileCleaner_CLI.ps1"

Starten in: C:\Pfad\Zum Skript\ (Zwingend erforderlich für die Pfadauflösung der Module!)

📈 Reporting

Nach jedem Durchlauf (GUI oder CLI) generiert das Tool im Ordner Logs/ einen zeitgestempelten Sitzungsordner. 

Dieser enthält:

-CSV-Rohdaten: Detaillierte Liste aller verarbeiteten Objekte inkl. Status.

-HTML-Report: Grafische Aufbereitung der Ergebnisse für das Monitoring.

-Log-File: Technisches Protokoll des Durchlaufs.

⚖️ Lizenz & Urheberschutz:

Dieses Projekt ist unter der GNU GPLv3 lizenziert. Dies stellt sicher, dass der Code offen bleibt, Verbesserungen geteilt werden müssen und mein Urheberrecht als Entwickler gewahrt bleibt.

Stand: 30.01.2026 (v1.0 Meilenstein erreicht)

Copyright 2026 Norman Thiel
