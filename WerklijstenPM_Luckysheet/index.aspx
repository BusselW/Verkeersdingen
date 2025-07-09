<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="Upload en bekijk Excel bestanden met Luckysheet - moderne spreadsheet viewer">
    <title>Verkeersdingen Werklijst - Luckysheet Viewer</title>
    <link href="styles.css" rel="stylesheet">
    <link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>📊</text></svg>">
    
    <!-- Luckysheet CSS -->
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/css/pluginsCss.css' />
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/plugins.css' />
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/css/luckysheet.css' />
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/assets/iconfont/iconfont.css' />
    
    <!-- External Dependencies -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/js/plugin.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/luckysheet.umd.js"></script>
</head>

<body>
    <div class="page-wrapper">
        <div class="container">
            <header class="header">
                <h1 class="title">Verkeersdingen Werklijst</h1>
                <p class="description">
                    Geavanceerde Excel viewer met Luckysheet voor professionele werklijst beheer
                </p>
                <div class="controls">
                    <button id="loadBtn" class="export-btn">Herlaad Data</button>
                    <button id="exportBtn" class="export-btn">Export Excel</button>
                </div>
            </header>
            
            <div class="stats">
                <span id="dataInfo">Rijen: 0, Kolommen: 0</span>
                <span id="sheetInfo">Sheet: Sheet1</span>
            </div>
            
            <div id="loading" class="loading-indicator">Luckysheet wordt geladen...</div>
            <div id="error" class="error-message" style="display: none;"></div>
            
            <div id="luckysheet" style="height: 500px; width: 100%;"></div>
        </div>
    </div>
    <script src="scripts.js"></script>
</body>
</html>
