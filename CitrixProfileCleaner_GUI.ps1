<#
.SYNOPSIS
    Citrix Profile Cleaner GUI 2026
.AUTHOR
    Norman Thiel
.VERSION
    1.0 (Meilenstein 30.01.2026)
.DESCRIPTION
    Grafische Benutzeroberfläche für die ProfileCleanupEngine. 
    Nutzt WPF für das UI-Rendering und PowerShell-Runspaces für die non-blocking Ausführung.
.FEATURES
    - Dynamisches Laden von Job-Definitionen aus JSON-Dateien.
    - Multithreading: GUI bleibt während des Löschvorgangs reaktionsfähig.
    - Echtzeit-Log-Streaming mit Farbkategorisierung (RichTextBox).
    - Session-Management: Erstellt pro Durchlauf einen Zeitstempel-Ordner für Logs, CSV und HTML.
    - Dry-Run Option zur sicheren Simulation von Löschvorgängen.
.DEPENDENCIES
    - ProfileCleanupEngine.psm1 (Kern-Logik)
    - Merge-ProfileCleanerSessionCSVs.psm1 (Reporting)
    - Admin-Rechte (für Dateizugriffe und Besitzübernahmen)
#>

# --- Pfadkonfiguration ---
# Ermittlung des Skript-Verzeichnisses für relative Pfadreferenzen
$projRoot = $PSScriptRoot
if (-not $projRoot) { $projRoot = Get-Location }

# Definition der zentralen Dateipfade für Module, Konfigurationen und Ausgaben
$mPath = Join-Path $projRoot "Modules\ProfileCleanupEngine.psm1"
$rMod  = Join-Path $projRoot "Modules\Merge-ProfileCleanerSessionCSVs.psm1"
$jPath = Join-Path $projRoot "Jobs"
$lPath = Join-Path $projRoot "Logs"

# --- GUI START LOGIK ---
# Erzwingt den Start in einem Single-Threaded Apartment (STA), notwendig für WPF
if ($args[0] -ne "--gui-run") {
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "$PSCommandPath" "--gui-run"
    exit
}

# --- GUI CODE BLOCK (WPF & Event Handling) ---
# Laden der benötigten .NET Assemblies für das User Interface
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Hauptfenster-Definition
$win = New-Object System.Windows.Window -Property @{ 
    Title = "Citrix Profile Cleaner 2026 - Professional"; 
    Width = 600; Height = 800; 
    Background = "#F0F0F0";
    WindowStartupLocation = "CenterScreen"
}

# Layout-Struktur: Unterteilung in Kopfbereich (Auto-Höhe) und Log-Bereich (Rest)
$grid = New-Object System.Windows.Controls.Grid
$win.Content = $grid
[void]$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height="Auto"}))
[void]$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height="*" }))

# Container für Steuerelemente im oberen Bereich
$top = New-Object System.Windows.Controls.StackPanel -Property @{ Margin=15 }
[System.Windows.Controls.Grid]::SetRow($top, 0)
[void]$grid.Children.Add($top)

# Kopfzeile für die Job-Auswahl mit Steuerungs-Buttons
$pnlJobHeader = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation="Horizontal"; Margin="0,0,0,5" }
$lblJobs = New-Object System.Windows.Controls.TextBlock -Property @{ Text="Verfügbare Cleanup-Jobs:"; FontWeight="Bold"; Width=200 }
$btnAll = New-Object System.Windows.Controls.Button -Property @{ Content="Alle anwählen"; Width=100; Height=20; FontSize=10; Margin="0,0,5,0" }
$btnNone = New-Object System.Windows.Controls.Button -Property @{ Content="Alle abwählen"; Width=100; Height=20; FontSize=10 }

[void]$pnlJobHeader.Children.Add($lblJobs)
[void]$pnlJobHeader.Children.Add($btnAll)
[void]$pnlJobHeader.Children.Add($btnNone)
[void]$top.Children.Add($pnlJobHeader)

# Scrollbereich für die dynamisch geladenen Jobs
$jobPanel = New-Object System.Windows.Controls.StackPanel
$scroll = New-Object System.Windows.Controls.ScrollViewer -Property @{ 
    Height=150; Content=$jobPanel; Background="White"; BorderBrush="#CCC"; BorderThickness="1"; VerticalScrollBarVisibility="Auto" 
}
[void]$top.Children.Add($scroll)

# Dynamisches Laden der Job-Definitionen aus dem JSON-Verzeichnis
if (Test-Path $jPath) {
    Get-ChildItem $jPath -Filter *.json | ForEach-Object {
        try {
            $j = Get-Content $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $cb = New-Object System.Windows.Controls.CheckBox -Property @{ 
                Content = $j.Label; Tag = $j; IsChecked = ($j.Enabled -eq $true); Margin = "5,2,5,2" 
            }
            [void]$jobPanel.Children.Add($cb)
        } catch { }
    }
}

# Bereich für Filter-Optionen und den Dry-Run Modus
$pnlCtrl = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation="Horizontal"; Margin="0,15,0,10" }
$chkI = New-Object System.Windows.Controls.CheckBox -Property @{ Content="INFOS"; IsChecked=$true; Margin="0,0,15,0" }
$chkW = New-Object System.Windows.Controls.CheckBox -Property @{ Content="WARNUNGEN"; IsChecked=$true; Margin="0,0,15,0" }
$chkE = New-Object System.Windows.Controls.CheckBox -Property @{ Content="FEHLER"; IsChecked=$true; Margin="0,0,25,0" }
$dry = New-Object System.Windows.Controls.CheckBox -Property @{ Content="DRY-RUN (Simulation)"; IsChecked=$true; FontWeight="Bold"; Foreground="#E67E22" }

foreach($c in @($chkI, $chkW, $chkE, $dry)) { [void]$pnlCtrl.Children.Add($c) }
[void]$top.Children.Add($pnlCtrl)

# Aktions-Leiste (Statusanzeige und Haupt-Buttons)
$pnlActions = New-Object System.Windows.Controls.Grid
[void]$pnlActions.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width="*"}))
[void]$pnlActions.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width="Auto"}))

$lblStatus = New-Object System.Windows.Controls.TextBlock -Property @{ Text="Status: Bereit"; FontWeight="Bold"; VerticalAlignment="Center" }
$pnlButtons = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation="Horizontal" }
$btnLogs = New-Object System.Windows.Controls.Button -Property @{ Content="Logs öffnen"; Width=100; Height=35; Margin="0,0,10,0"; Background="#6c757d"; Foreground="White" }
$btnStart = New-Object System.Windows.Controls.Button -Property @{ Content="START"; Width=120; Height=35; Background="#28a745"; Foreground="White"; FontWeight="Bold" }
$btnStop = New-Object System.Windows.Controls.Button -Property @{ Content="STOPP"; Width=100; Height=35; Background="#dc3545"; Foreground="White"; IsEnabled=$false; Margin="10,0,0,0" }

[void]$pnlButtons.Children.Add($btnLogs)
[void]$pnlButtons.Children.Add($btnStart)
[void]$pnlButtons.Children.Add($btnStop)
[void][System.Windows.Controls.Grid]::SetColumn($lblStatus, 0)
[void][System.Windows.Controls.Grid]::SetColumn($pnlButtons, 1)
[void]$pnlActions.Children.Add($lblStatus)
[void]$pnlActions.Children.Add($pnlButtons)
[void]$top.Children.Add($pnlActions)

# Zentrales Log-Fenster (RichTextBox) mit dunklem Design für Konsolen-Feeling
$logBox = New-Object System.Windows.Controls.RichTextBox -Property @{ 
    IsReadOnly=$true; Background="#1E1E1E"; FontFamily="Consolas"; Foreground="White"; VerticalScrollBarVisibility="Auto"
}
[void][System.Windows.Controls.Grid]::SetRow($logBox, 1)
[void]$grid.Children.Add($logBox)

# Synchronisierte Hashtable: Brücke zwischen GUI-Thread und Runspace-Hintergrundthread
$Sync = [hashtable]::Synchronized(@{ 
    CancelRequested = $false; 
    Finished        = $false; 
    ReportData      = New-Object System.Collections.ArrayList; 
    CurrentJob      = "Bereit"; 
    Window          = $win;
    SessionPath     = "";
    LogFile         = "" 
})

# Event-Handler für UI-Interaktionen
$btnAll.Add_Click({ foreach($cb in $jobPanel.Children) { $cb.IsChecked = $true } })
$btnNone.Add_Click({ foreach($cb in $jobPanel.Children) { $cb.IsChecked = $false } })
$btnLogs.Add_Click({ if(Test-Path $lPath) { Invoke-Item $lPath } })

# Haupt-Laufzeit-Logik beim Klicken auf START
$btnStart.Add_Click({
    # Ermittlung der aktuell ausgewählten Jobs
    $selected = @(); foreach($c in $jobPanel.Children) { if($c.IsChecked) { $selected += $c.Tag } }
    if ($selected.Count -eq 0) { return }

    # UI-Status auf "Beschäftigt" setzen
    $Sync.CancelRequested = $false; $Sync.Finished = $false; $Sync.ReportData.Clear()
    $btnStart.IsEnabled = $false; $btnStop.IsEnabled = $true; 
    $dry.IsEnabled = $false;
    $logBox.Document.Blocks.Clear()

    # Erstellung des Sitzungs-Verzeichnisses für detaillierte Protokolle
    $sessionTimestamp = Get-Date -Format 'yyyyMMdd_HHmm'
    $Sync.SessionPath = Join-Path $lPath "Session_$sessionTimestamp"
    if (!(Test-Path $Sync.SessionPath)) { New-Item $Sync.SessionPath -ItemType Directory -Force | Out-Null }
    
    $Sync.LogFile = Join-Path $Sync.SessionPath "Cleanup_Details.log"
    "--- GUI SESSION START: $(Get-Date) ---" | Out-File $Sync.LogFile -Encoding UTF8

    # Initialisierung des Hintergrund-Arbeitsthreads (Runspace)
    $rs = [runspacefactory]::CreateRunspace(); $rs.ApartmentState = "STA"; $rs.Open()
    $ps = [powershell]::Create().AddScript({
        param($jobs, $isDry, $mod, $Sync, $ci, $cw, $ce, $log)
        
        # Interne Logging-Funktion innerhalb des Runspaces zur direkten Manipulation der GUI-RichTextBox
        function global:Write-Log-Internal {
            param($m, $lvl)
            # Filter für unwichtige Systemmeldungen
            if ($m -like "*now owned by user*" -or $m -like "*Successfully processed*") { return }
            # Filterung basierend auf den Checkboxen in der GUI
            if (($lvl -eq "INFO" -or $lvl -eq "PASSIV" -or $lvl -eq "SUCCESS") -and -not $ci.IsChecked) { return }
            if ($lvl -eq "WARN" -and -not $cw.IsChecked) { return }
            if ($lvl -eq "ERROR" -and -not $ce.IsChecked) { return }
            
            # Icon-Zuordnung und Zeitstempel
            $sym = switch($lvl) { "SUCCESS"{"[OK] "} "PASSIV"{"[...] "} "WARN"{"[!] "} "ERROR"{"[!] "} default{""} }
            $time = Get-Date -Format "HH:mm:ss"
            $p = New-Object System.Windows.Documents.Paragraph -Property @{Margin="0"}
            $r = New-Object System.Windows.Documents.Run("[$time] $sym$m")
            
            # Farbliche Kennzeichnung basierend auf Log-Level
            switch ($lvl) {
                "SUCCESS" { $r.Foreground="LightGreen"; $r.FontWeight="Bold" }
                "WARN" { $r.Foreground="Yellow" }
                "ERROR" { $r.Foreground="OrangeRed"; $r.FontWeight="Bold" }
                "PASSIV" { $r.Foreground="Gray" }
                default { $r.Foreground="White" }
            }
            # Thread-sicheres Hinzufügen der Nachricht zum UI
            [void]$p.Inlines.Add($r); [void]$log.Document.Blocks.Add($p); $log.ScrollToEnd()
        }

        # Import der Cleanup-Kernlogik und sequenzielle Abarbeitung der Jobs
        Import-Module $mod -Force
        foreach($j in $jobs) { 
            if($Sync.CancelRequested){break}
            Invoke-ProfileCleanupJob -Job $j -DryRun:$isDry -Sync $Sync 
        }
        $Sync.Finished = $true
    }).AddArgument($selected).AddArgument($dry.IsChecked).AddArgument($mPath).AddArgument($Sync).AddArgument($chkI).AddArgument($chkW).AddArgument($chkE).AddArgument($logBox)
    
    $ps.Runspace = $rs; [void]$ps.BeginInvoke()

    # GUI-Timer zur periodischen Überwachung des Hintergrund-Status (Jede Sekunde)
    $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval=[TimeSpan]::FromSeconds(1) }
    $timer.Add_Tick({
        $lblStatus.Text = "Status: $($Sync.CurrentJob)"
        # Abschluss-Logik wenn der Hintergrundprozess beendet wurde
        if ($Sync.Finished) {
            $this.Stop()
            if ($Sync.ReportData.Count -gt 0 -and $Sync.SessionPath -ne "") {
                $auditUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                
                # Report-Daten finalisieren und Details aus Engine-Objekten mappen
                $finalData = foreach ($row in $Sync.ReportData) {
                    $details = if ($row.AuditDetails) { $row.AuditDetails } elseif ($row.Message) { $row.Message } else { "Aktion ausgeführt" }
                    
                    # Generierung des finalen Objekts für den CSV/HTML Export
                    $row | Select-Object *, 
                        @{Name='RunBy'; Expression={"$auditUser (GUI)"}}, 
                        @{Name='AuditDetails'; Expression={$details}} -ExcludeProperty RunBy, AuditDetails, Message
                }
                
                # Datenexport in CSV
                $csvPath = Join-Path $Sync.SessionPath "Cleanup_Data.csv"
                $finalData | Export-Csv $csvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
                
                # Aufruf des Reporting-Moduls für HTML-Generierung und automatische Anzeige
                if (Test-Path $rMod) { 
                    Import-Module $rMod -Force
                    $cmd = Get-Command Merge-ProfileCleanerSessionCSVs -ErrorAction SilentlyContinue
                    if ($cmd) {
                         & $cmd $Sync.SessionPath 
                         if (Test-Path $Sync.SessionPath) { Invoke-Item $Sync.SessionPath }
                    }
                }
            }
            # UI zurück auf Ursprung setzen
            $btnStart.IsEnabled = $true; $btnStop.IsEnabled = $false; 
            $dry.IsEnabled = $true; 
            $lblStatus.Text = "Status: Fertig"
        }
    })
    $timer.Start()
})

# Stopp-Funktionalität: Setzt nur das Abbruch-Flag, die Engine bricht dann sauber ab
$btnStop.Add_Click({ $Sync.CancelRequested = $true; $btnStop.IsEnabled = $false })

# Starten der GUI-Schleife
[void]$win.ShowDialog()