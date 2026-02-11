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
        [string]$SessionPath 
    )

    if (-not (Test-Path $SessionPath)) { return }

    $csvFile = Join-Path $SessionPath "Cleanup_Data.csv"
    $logFileRelative = "Cleanup_Details.log"
    
    # --- DATENIMPORT ---
    if (-not (Test-Path $csvFile)) {
        $allCsvs = Get-ChildItem -Path $SessionPath -Filter "*.csv" -File
        if ($allCsvs.Count -eq 0) { return }
        $data = $allCsvs | ForEach-Object { Import-Csv $_.FullName -Delimiter ";" }
    } else {
        $data = Import-Csv $csvFile -Delimiter ";"
    }

    # --- DATENAUFBEREITUNG ---
    foreach($row in $data) {
        if ($row.MB) { $row.MB = [double]($row.MB -replace ',', '.') }
        if (-not $row.PSObject.Properties['RunBy']) { 
            Add-Member -InputObject $row -NoteProperty "RunBy" -Value "N/A" 
        }
    }

    # --- BERECHNUNG DER KENNZAHLEN (DYNAMISCH) ---
    $relevantData = $data | Where-Object { $_.Status -eq "Erfolgreich gelöscht" -or $_.Status -eq "SIMULATION" }
    $totalMB = ($relevantData | Measure-Object -Property MB -Sum).Sum
    
    $isSimulationMode = $data.Status -contains "SIMULATION"
    $labelEinsparung = if ($isSimulationMode) { "Mögliche Einsparung" } else { "Einsparung" }
    $statColor = if ($isSimulationMode) { "#f59e0b" } else { "#107c10" }

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
            --neutral-bg: #f3f2f1; --neutral-text: #605e5c;
        }
        body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 40px; }
        .container { max-width: 1700px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); overflow-x: auto; }
        .header-flex { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #edebe9; padding-bottom: 20px; margin-bottom: 20px; }
        h2 { margin: 0; color: var(--primary); font-weight: 300; font-size: 26px; }
        .btn-log { 
            background: var(--primary); color: white !important; text-decoration: none; padding: 10px 20px; 
            border-radius: 4px; font-size: 13px; font-weight: 600; transition: all 0.2s ease;
            display: inline-flex; align-items: center; gap: 10px; border: none; cursor: pointer;
        }
        .btn-log:hover { background: var(--primary-hover); transform: translateY(-1px); box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
        .summary-card { display: flex; gap: 40px; margin: 20px 0 30px 0; background: #faf9f8; padding: 25px; border-radius: 6px; border-left: 4px solid var(--primary); }
        .stat-item { display: flex; flex-direction: column; }
        .stat-label { font-size: 11px; text-transform: uppercase; color: #605e5c; letter-spacing: 0.8px; margin-bottom: 5px; }
        .stat-value { font-size: 22px; font-weight: 600; color: var(--text); }
        
        /* TABLE DYNAMISCH ANPASSEN */
        table { width: 100%; border-collapse: collapse; table-layout: auto; min-width: 100%; }
        th { 
            background: #faf9f8; text-align: left; padding: 15px 12px; font-size: 13px; 
            cursor: pointer; border-bottom: 2px solid #edebe9; position: sticky; top: 0; z-index: 10;
            white-space: nowrap; user-select: none; position: relative;
        }
        .resizer {
            position: absolute; right: 0; top: 0; width: 6px; height: 100%;
            cursor: col-resize; z-index: 11;
        }
        .resizer:hover { background: var(--primary); opacity: 0.5; }

        td { padding: 12px; font-size: 12px; border-bottom: 1px solid #edebe9; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        tr:hover { background-color: #fcfcfc; }
        .status { padding: 4px 12px; border-radius: 15px; font-size: 11px; font-weight: 600; display: inline-block; }
        .status-Erfolgreich-gelöscht, .status-ERFOLG { background: var(--success-bg); color: var(--success-text); }
        .status-SIMULATION { background: var(--warn-bg); color: #9a6700; border: 1px solid #ffd33d; }
        .status-FEHLER { background: var(--error-bg); color: var(--error-text); }
        .status-GEPRÜFT { background: var(--neutral-bg); color: var(--neutral-text); font-weight: normal; }
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
            <div class="stat-item"><div class="stat-label">Analysierte Objekte</div><div class="stat-value">$($data.Count)</div></div>
            <div class="stat-item">
                <div class="stat-label">$labelEinsparung</div>
                <div class="stat-value" style="color:$statColor">$displayTotal</div>
            </div>
        </div>
        
        <table id="reportTable">
            <thead>
                <tr>
                    <th>Job</th>
                    <th>Pfad</th>
                    <th>Größe</th>
                    <th>Alter</th>
                    <th>Status</th>
                    <th>Ausgeführt von</th>
                </tr>
            </thead>
            <tbody>
"@

    # --- TABELLENZEILEN GENERIEREN ---
    $htmlRows = $data | ForEach-Object {
        $valMB = [double]$_.MB
        $valAge = if ($_.Alter -as [int]) { [int]$_.Alter } else { 0 }
        $displayAge = if ($_.Alter -eq "N/A") { "N/A" } else { "$valAge Tage" }
        $displaySize = if ($valMB -ge 1024) { "$([math]::Round($valMB / 1024, 2)) GB" } else { "$valMB MB" }
        $statusClass = "status-$($_.Status.Replace(' ', '-'))"
        
        "                <tr>
                    <td title='$($_.Job)'>$($_.Job)</td>
                    <td title='$($_.Pfad)'>$($_.Pfad)</td>
                    <td data-sort='$valMB'>$displaySize</td>
                    <td data-sort='$valAge'>$displayAge</td>
                    <td><span class='status $statusClass'>$($_.Status)</span></td>
                    <td title='$($_.RunBy)'>$($_.RunBy)</td>
                </tr>"
    }

    $htmlFoot = @"
            </tbody>
        </table>
    </div>
    <script>
        // --- PERFORMANCE SORTIERUNG (DOM-DETACH) ---
        const getCellValue = (tr, idx) => {
            const td = tr.children[idx];
            return td.getAttribute('data-sort') || td.innerText || td.textContent;
        };

        const comparer = (idx, asc) => (a, b) => ((v1, v2) => 
            v1 !== '' && v2 !== '' && !isNaN(v1) && !isNaN(v2) ? v1 - v2 : v1.toString().localeCompare(v2)
            )(getCellValue(asc ? a : b, idx), getCellValue(asc ? b : a, idx));

        document.querySelectorAll('th').forEach(th => th.addEventListener('click', function(e) {
            if (e.target.classList.contains('resizer')) return;
            
            const table = th.closest('table');
            const tbody = table.querySelector('tbody');
            document.body.style.cursor = 'wait';
            
            const parent = tbody.parentNode;
            parent.removeChild(tbody); // Detach für Performance

            const rows = Array.from(tbody.querySelectorAll('tr'));
            const index = Array.from(th.parentNode.children).indexOf(th);
            this.asc = !this.asc;

            rows.sort(comparer(index, this.asc)).forEach(tr => tbody.appendChild(tr));

            parent.appendChild(tbody);
            document.body.style.cursor = 'default';
        }));

        // --- DYNAMISCHER RESIZER (HYBRID) ---
        document.querySelectorAll('th').forEach(th => {
            const resizer = document.createElement('div');
            resizer.classList.add('resizer');
            th.appendChild(resizer);

            resizer.addEventListener('mousedown', function(e) {
                e.stopPropagation();
                e.preventDefault();
                
                const startX = e.pageX;
                const startWidth = th.offsetWidth;
                const table = th.closest('table');

                // Fixiere Layout erst beim ersten Resize
                if (table.style.tableLayout !== 'fixed') {
                    table.querySelectorAll('th').forEach(h => h.style.width = h.offsetWidth + 'px');
                    table.style.tableLayout = 'fixed';
                }

                const onMouseMove = (e) => {
                    const newWidth = (startWidth + (e.pageX - startX)) + 'px';
                    th.style.width = newWidth;
                    th.style.minWidth = newWidth;
                };

                const onMouseUp = () => {
                    document.removeEventListener('mousemove', onMouseMove);
                    document.removeEventListener('mouseup', onMouseUp);
                };

                document.addEventListener('mousemove', onMouseMove);
                document.addEventListener('mouseup', onMouseUp);
            });
        });
    </script>
</body>
</html>
"@

    $htmlHead + ($htmlRows -join "`n") + $htmlFoot | Out-File $reportFile -Encoding UTF8
}

Export-ModuleMember -Function Merge-ProfileCleanerSessionCSVs