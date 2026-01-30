<#
.SYNOPSIS
    Profile Cleaner Report Merger 2026
.AUTHOR
    Norman Thiel
.VERSION
    1.0 (Meilenstein 30.01.2026)
.DESCRIPTION
    Aggregiert CSV-Daten einer Reinigungssitzung zu einem interaktiven HTML-Bericht.
    Enthält eine Zusammenfassung der Einsparungen und detaillierte Status-Informationen.
.FEATURES
    - Fest integrierter "Log öffnen"-Button im Header für direkten Zugriff auf Cleanup_Details.log.
    - Dynamische Einheiten-Umrechnung (MB zu GB) für die Gesamteinsparung.
    - Flat Design UI mit CSS-Varianten für Status-Farbcodierung.
    - JavaScript-basierte Tabellensortierung direkt im Browser.
    - Robustes Daten-Mapping (Punkt/Komma-Korrektur bei Größenangaben).
.NOTES
    Abhängigkeiten: Cleanup_Data.csv im SessionPath.
#>

function Merge-ProfileCleanerSessionCSVs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$SessionPath # Zielverzeichnis der Sitzung
    )

    # Validierung: Pfad muss existieren
    if (-not (Test-Path $SessionPath)) { return }

    # Dateidefinitionen
    $csvFile = Join-Path $SessionPath "Cleanup_Data.csv"
    $logFileRelative = "Cleanup_Details.log"
    
    # --- DATENIMPORT ---
    # Fallback-Logik: Falls die Haupt-CSV fehlt, werden alle Einzel-CSVs im Ordner aggregiert
    if (-not (Test-Path $csvFile)) {
        $allCsvs = Get-ChildItem -Path $SessionPath -Filter "*.csv" -File
        if ($allCsvs.Count -eq 0) { return }
        $data = $allCsvs | ForEach-Object { Import-Csv $_.FullName -Delimiter ";" }
    } else {
        $data = Import-Csv $csvFile -Delimiter ";"
    }

    # --- DATENAUFBEREITUNG ---
    foreach($row in $data) {
        # Konvertiert Komma in Punkt für mathematische Operationen in PowerShell
        if ($row.MB) { $row.MB = [double]($row.MB -replace ',', '.') }
        
        # Sicherstellen, dass Audit-Felder vorhanden sind (verhindert HTML-Lücken)
        if (-not $row.PSObject.Properties['RunBy']) { 
            Add-Member -InputObject $row -NoteProperty "RunBy" -Value "N/A" 
        }
    }

    # Berechnung der Kennzahlen für den Header
    $totalMB = ($data | Measure-Object -Property MB -Sum).Sum
    $timeStamp = Get-Date -Format "dd.MM.yyyy HH:mm"
    $displayTotal = if ($totalMB -ge 1024) { 
        "$([math]::Round($totalMB / 1024, 2)) GB" 
    } else { 
        "$([math]::Round($totalMB, 2)) MB" 
    }

    $reportFile = Join-Path $SessionPath "Cleanup_Report.html"

    # --- HTML STRUKTUR & CSS ---
    $htmlHead = @"
<!DOCTYPE html>
<html lang='de'>
<head>
    <meta charset='UTF-8'>
    <title>Cleanup Report - $timeStamp</title>
    <style>
        :root {
            --primary: #0078d4; --primary-hover: #005a9e; --bg: #f3f2f1; --text: #323130;
            --success-bg: #dff6dd; --success-text: #107c10;
            --warn-bg: #fff4ce; --warn-text: #797775;
            --error-bg: #fde7e9; --error-text: #a4262c;
        }
        body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 40px; }
        .container { max-width: 1600px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }
        
        /* Header & Navigation */
        .header-flex { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #edebe9; padding-bottom: 20px; margin-bottom: 20px; }
        h2 { margin: 0; color: var(--primary); font-weight: 300; font-size: 26px; }
        
        /* Action Button */
        .btn-log { 
            background: var(--primary); color: white !important; text-decoration: none; padding: 10px 20px; 
            border-radius: 4px; font-size: 13px; font-weight: 600; transition: all 0.2s ease;
            display: inline-flex; align-items: center; gap: 10px; border: none; cursor: pointer;
        }
        .btn-log:hover { background: var(--primary-hover); transform: translateY(-1px); box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
        .btn-log:active { transform: translateY(0); }

        /* Infokarten */
        .summary-card { display: flex; gap: 40px; margin: 20px 0 30px 0; background: #faf9f8; padding: 25px; border-radius: 6px; border-left: 4px solid var(--primary); }
        .stat-item { display: flex; flex-direction: column; }
        .stat-label { font-size: 11px; text-transform: uppercase; color: #605e5c; letter-spacing: 0.8px; margin-bottom: 5px; }
        .stat-value { font-size: 22px; font-weight: 600; color: var(--text); }
        
        /* Tabellen Styling */
        table { width: 100%; border-collapse: collapse; table-layout: fixed; }
        th { background: #faf9f8; text-align: left; padding: 15px 12px; font-size: 13px; cursor: pointer; border-bottom: 2px solid #edebe9; position: sticky; top: 0; z-index: 10; }
        th:hover { background: #f3f2f1; }
        td { padding: 12px; font-size: 12px; border-bottom: 1px solid #edebe9; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        tr:hover { background-color: #fcfcfc; }
        
        /* Status Badges */
        .status { padding: 4px 12px; border-radius: 15px; font-size: 11px; font-weight: 600; display: inline-block; }
        .status-Erfolgreich-gelöscht, .status-ERFOLG { background: var(--success-bg); color: var(--success-text); }
        .status-SIMULATION { background: var(--warn-bg); color: var(--warn-text); }
        .status-FEHLER { background: var(--error-bg); color: var(--error-text); }
        .status-GEPRÜFT { background: #eee; color: #666; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header-flex">
            <h2>Citrix ProfileCleaner Report</h2>
            <a href="$logFileRelative" target="_blank" class="btn-log">
                <span style="font-size: 16px;">📋</span> Vollständiges Log einsehen
            </a>
        </div>

        <div class="summary-card">
            <div class="stat-item"><div class="stat-label">Zeitstempel</div><div class="stat-value">$timeStamp</div></div>
            <div class="stat-item"><div class="stat-label">Analysierte Profile</div><div class="stat-value">$($data.Count)</div></div>
            <div class="stat-item"><div class="stat-label">Einsparung</div><div class="stat-value">$displayTotal</div></div>
        </div>
        
        <table id="reportTable">
            <thead>
                <tr>
                    <th style="width:12%">Job</th>
                    <th style="width:38%">Pfad</th>
                    <th style="width:10%">Größe</th>
                    <th style="width:8%">Alter</th>
                    <th style="width:17%">Status</th>
                    <th style="width:15%">Ausgeführt von</th>
                </tr>
            </thead>
            <tbody>
"@

    # --- TABELLENZEILEN GENERIEREN ---
    $htmlRows = $data | ForEach-Object {
        $valMB = [double]$_.MB
        $valAge = [int]$_.Alter
        $displaySize = if ($valMB -ge 1024) { "$([math]::Round($valMB / 1024, 2)) GB" } else { "$valMB MB" }
        $statusClass = "status-$($_.Status.Replace(' ', '-'))"
        
        "                <tr>
                    <td title='$($_.Job)'>$($_.Job)</td>
                    <td title='$($_.Pfad)'>$($_.Pfad)</td>
                    <td data-sort='$valMB'>$displaySize</td>
                    <td data-sort='$valAge'>$valAge Tage</td>
                    <td><span class='status $statusClass'>$($_.Status)</span></td>
                    <td title='$($_.RunBy)'>$($_.RunBy)</td>
                </tr>"
    }

    # --- FOOTER & JAVASCRIPT ---
    $htmlFoot = @"
            </tbody>
        </table>
    </div>
    <script>
        // Tabellensortierung
        const getCellValue = (tr, idx) => {
            const td = tr.children[idx];
            return td.getAttribute('data-sort') || td.innerText || td.textContent;
        };
        const comparer = (idx, asc) => (a, b) => ((v1, v2) => 
            v1 !== '' && v2 !== '' && !isNaN(v1) && !isNaN(v2) ? v1 - v2 : v1.toString().localeCompare(v2)
            )(getCellValue(asc ? a : b, idx), getCellValue(asc ? b : a, idx));

        document.querySelectorAll('th').forEach(th => th.addEventListener('click', function() {
            const table = th.closest('table');
            const tbody = table.querySelector('tbody');
            const index = Array.from(th.parentNode.children).indexOf(th);
            this.asc = !this.asc;
            Array.from(tbody.querySelectorAll('tr'))
                .sort(comparer(index, this.asc))
                .forEach(tr => tbody.appendChild(tr));
        }));
    </script>
</body>
</html>
"@

    # Report speichern
    $htmlHead + ($htmlRows -join "`n") + $htmlFoot | Out-File $reportFile -Encoding UTF8
}

Export-ModuleMember -Function Merge-ProfileCleanerSessionCSVs