<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="Upload en bekijk Excel bestanden met SheetJS - moderne spreadsheet viewer">
    <title>Verkeersdingen Werklijst - SheetJS Viewer</title>
    <link href="styles.css" rel="stylesheet">
    <link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>📊</text></svg>">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>

<body>
    <div class="page-wrapper">
        <div class="container">
            <header class="header">
                <h1 class="title">Verkeersdingen Werklijst</h1>
                <p class="description">
                    Betrouwbare Excel viewer met SheetJS voor nauwkeurige werklijst weergave
                </p>
                <div class="controls">
                    <a href="https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1" 
                       class="download-icon" target="_blank" rel="noopener noreferrer" title="Download Excel bestand"></a>
                    <button id="loadBtn" class="export-btn">Herlaad Data</button>
                    <button id="exportBtn" class="export-btn">Export Excel</button>
                </div>
            </header>
            
            <div class="stats">
                <span id="dataInfo">Rijen: 0, Kolommen: 0</span>
                <span id="sheetInfo">Sheet: Sheet1</span>
            </div>
            
            <div id="loading" class="loading-indicator">Excel data wordt geladen...</div>
            <div id="error" class="error-message" style="display: none;"></div>
            
            <div id="spreadsheet-container">
                <div class="top-scrollbar" id="topScrollbar">
                    <div></div>
                </div>
                <div class="table-container" id="tableContainer">
                    <table id="dataTable" class="data-table">
                        <thead id="tableHeader"></thead>
                        <tbody id="tableBody"></tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <script src="scripts.js"></script>
</body>
</html>
