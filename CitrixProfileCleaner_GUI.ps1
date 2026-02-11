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
$projRoot = $PSScriptRoot
if (-not $projRoot) { $projRoot = Get-Location }

# Zentrale Pfade zu Modulen, Jobs und Logs
$mPath = Join-Path $projRoot "Modules\ProfileCleanupEngine.psm1"
$rMod  = Join-Path $projRoot "Modules\Merge-ProfileCleanerSessionCSVs.psm1"
$jPath = Join-Path $projRoot "Jobs"
$lPath = Join-Path $projRoot "Logs"

# --- GUI START LOGIK (Fix für 'The filename or extension is too long') ---
# Wir prüfen, ob das Skript bereits im STA-Modus läuft.
if ($args[0] -ne "--gui-run") {
    # Wir übergeben nur den Pfad der Datei via -File (stabil gegen lange Skripte)
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "$PSCommandPath" "--gui-run"
    exit
}

# --- GUI CODE BLOCK (WPF & Event Handling) ---
# Laden der notwendigen .NET Framework Assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Hauptfenster Definition
$win = New-Object System.Windows.Window -Property @{ 
    Title = "Citrix Profile Cleaner 2026 - Professional"; 
    Width = 600; Height = 800; 
    Background = "#F0F0F0";
    WindowStartupLocation = "CenterScreen"
}

# Layout-Grid Definition
$grid = New-Object System.Windows.Controls.Grid
$win.Content = $grid
[void]$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height="Auto"}))
[void]$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height="*" }))

$top = New-Object System.Windows.Controls.StackPanel -Property @{ Margin=15 }
[System.Windows.Controls.Grid]::SetRow($top, 0)
[void]$grid.Children.Add($top)

# Sektion: Job Auswahl
$pnlJobHeader = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation="Horizontal"; Margin="0,0,0,5" }
$lblJobs = New-Object System.Windows.Controls.TextBlock -Property @{ Text="Verfügbare Cleanup-Jobs:"; FontWeight="Bold"; Width=200 }
$btnAll = New-Object System.Windows.Controls.Button -Property @{ Content="Alle anwählen"; Width=100; Height=20; FontSize=10; Margin="0,0,5,0" }
$btnNone = New-Object System.Windows.Controls.Button -Property @{ Content="Alle abwählen"; Width=100; Height=20; FontSize=10 }

[void]$pnlJobHeader.Children.Add($lblJobs)
[void]$pnlJobHeader.Children.Add($btnAll)
[void]$pnlJobHeader.Children.Add($btnNone)
[void]$top.Children.Add($pnlJobHeader)

# Scrollbarer Bereich für Checkboxen
$jobPanel = New-Object System.Windows.Controls.StackPanel
$scroll = New-Object System.Windows.Controls.ScrollViewer -Property @{ 
    Height=150; Content=$jobPanel; Background="White"; BorderBrush="#CCC"; BorderThickness="1"; VerticalScrollBarVisibility="Auto" 
}
[void]$top.Children.Add($scroll)

# Dynamischer Job-Import
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

# Sektion: Filter & Optionen
$pnlCtrl = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation="Horizontal"; Margin="0,15,0,10" }
$chkI = New-Object System.Windows.Controls.CheckBox -Property @{ Content="INFOS"; IsChecked=$true; Margin="0,0,15,0" }
$chkW = New-Object System.Windows.Controls.CheckBox -Property @{ Content="WARNUNGEN"; IsChecked=$true; Margin="0,0,15,0" }
$chkE = New-Object System.Windows.Controls.CheckBox -Property @{ Content="FEHLER"; IsChecked=$true; Margin="0,0,25,0" }
$dry = New-Object System.Windows.Controls.CheckBox -Property @{ Content="DRY-RUN (Simulation)"; IsChecked=$true; FontWeight="Bold"; Foreground="#E67E22" }

foreach($c in @($chkI, $chkW, $chkE, $dry)) { [void]$pnlCtrl.Children.Add($c) }
[void]$top.Children.Add($pnlCtrl)

# Sektion: Status & Aktions-Buttons
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

# Sektion: Log Ausgabe (RichTextBox)
$logBox = New-Object System.Windows.Controls.RichTextBox -Property @{ 
    IsReadOnly=$true; Background="#1E1E1E"; FontFamily="Consolas"; Foreground="White"; VerticalScrollBarVisibility="Auto"
}
[void][System.Windows.Controls.Grid]::SetRow($logBox, 1)
[void]$grid.Children.Add($logBox)

# Thread-Sichere Kommunikations-Hashtable
$Sync = [hashtable]::Synchronized(@{ 
    CancelRequested = $false; 
    Finished        = $false; 
    ReportData      = New-Object System.Collections.ArrayList; 
    CurrentJob      = "Bereit"; 
    Window          = $win;
    SessionPath     = "";
    LogFile         = "" 
})

# UI Event Handlers
$btnAll.Add_Click({ foreach($cb in $jobPanel.Children) { $cb.IsChecked = $true } })
$btnNone.Add_Click({ foreach($cb in $jobPanel.Children) { $cb.IsChecked = $false } })
$btnLogs.Add_Click({ if(Test-Path $lPath) { Invoke-Item $lPath } })

# Haupt-Logik beim Klick auf START
$btnStart.Add_Click({
    $selected = @(); foreach($c in $jobPanel.Children) { if($c.IsChecked) { $selected += $c.Tag } }
    if ($selected.Count -eq 0) { return }

    $Sync.CancelRequested = $false; $Sync.Finished = $false; $Sync.ReportData.Clear()
    $btnStart.IsEnabled = $false; $btnStop.IsEnabled = $true; 
    $dry.IsEnabled = $false;
    $logBox.Document.Blocks.Clear()

    $sessionTimestamp = Get-Date -Format 'yyyyMMdd_HHmm'
    $Sync.SessionPath = Join-Path $lPath "Session_$sessionTimestamp"
    if (!(Test-Path $Sync.SessionPath)) { New-Item $Sync.SessionPath -ItemType Directory -Force | Out-Null }
    
    $Sync.LogFile = Join-Path $Sync.SessionPath "Cleanup_Details.log"
    "--- GUI SESSION START: $(Get-Date) ---" | Out-File $Sync.LogFile -Encoding UTF8

    # Hintergrund-Thread (Runspace) initialisieren
    $rs = [runspacefactory]::CreateRunspace(); $rs.ApartmentState = "STA"; $rs.Open()
    $ps = [powershell]::Create().AddScript({
        param($jobs, $isDry, $mod, $Sync, $ci, $cw, $ce, $log)
        
        # Interne Log-Funktion
        function global:Write-Log-Internal {
            param($m, $lvl)
            if ($m -like "*now owned by user*" -or $m -like "*Successfully processed*") { return }
            if (($lvl -eq "INFO" -or $lvl -eq "PASSIV" -or $lvl -eq "SUCCESS") -and -not $ci.IsChecked) { return }
            if ($lvl -eq "WARN" -and -not $cw.IsChecked) { return }
            if ($lvl -eq "ERROR" -and -not $ce.IsChecked) { return }
            
            $sym = switch($lvl) { "SUCCESS"{"[OK] "} "PASSIV"{"[...] "} "WARN"{"[!] "} "ERROR"{"[!] "} default{""} }
            $time = Get-Date -Format "HH:mm:ss"
            $p = New-Object System.Windows.Documents.Paragraph -Property @{Margin="0"}
            $r = New-Object System.Windows.Documents.Run("[$time] $sym$m")
            
            switch ($lvl) {
                "SUCCESS" { $r.Foreground="LightGreen"; $r.FontWeight="Bold" }
                "WARN" { $r.Foreground="Yellow" }
                "ERROR" { $r.Foreground="OrangeRed"; $r.FontWeight="Bold" }
                "PASSIV" { $r.Foreground="Gray" }
                default { $r.Foreground="White" }
            }
            [void]$p.Inlines.Add($r); [void]$log.Document.Blocks.Add($p); $log.ScrollToEnd()
        }

        Import-Module $mod -Force
        foreach($j in $jobs) { 
            if($Sync.CancelRequested){break}
            Invoke-ProfileCleanupJob -Job $j -DryRun:$isDry -Sync $Sync 
        }
        $Sync.Finished = $true
    }).AddArgument($selected).AddArgument($dry.IsChecked).AddArgument($mPath).AddArgument($Sync).AddArgument($chkI).AddArgument($chkW).AddArgument($chkE).AddArgument($logBox)
    
    $ps.Runspace = $rs; [void]$ps.BeginInvoke()

    # UI-Timer zur Überwachung
    $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval=[TimeSpan]::FromSeconds(1) }
    $timer.Add_Tick({
        $lblStatus.Text = "Status: $($Sync.CurrentJob)"
        if ($Sync.Finished) {
            $this.Stop()
            if ($Sync.ReportData.Count -gt 0 -and $Sync.SessionPath -ne "") {
                $auditUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                foreach ($row in $Sync.ReportData) {
                    if ($null -eq $row.RunBy -or [string]::IsNullOrWhiteSpace($row.RunBy)) { 
                        $row.RunBy = "$auditUser (GUI)"
                    }
                    if ($null -eq $row.AuditDetails) {
                        $row | Add-Member -MemberType NoteProperty -Name "AuditDetails" -Value "" -Force -ErrorAction SilentlyContinue
                    }
                }
                
                $csvPath = Join-Path $Sync.SessionPath "Cleanup_Data.csv"
                $Sync.ReportData | Export-Csv $csvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
                
                if (Test-Path $rMod) { 
                    Import-Module $rMod -Force
                    Merge-ProfileCleanerSessionCSVs -SessionPath $Sync.SessionPath 
                    Invoke-Item $Sync.SessionPath 
                }
            }
            $btnStart.IsEnabled = $true; $btnStop.IsEnabled = $false; 
            $dry.IsEnabled = $true; 
            $lblStatus.Text = "Status: Fertig"
        }
    })
    $timer.Start()
})

# Stopp-Signal senden
$btnStop.Add_Click({ $Sync.CancelRequested = $true; $btnStop.IsEnabled = $false })

# GUI Start
[void]$win.ShowDialog()