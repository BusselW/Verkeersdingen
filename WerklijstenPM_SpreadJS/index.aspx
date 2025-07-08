<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verkeersborden Werklijst - SpreadJS</title>
    <link href="styles.css" rel="stylesheet">
    <script src="https://cdn.grapecity.com/spreadjs/hosted/spreadjs.runtime.all.min.js" type="text/javascript"></script>
    <link rel="stylesheet" type="text/css" href="https://cdn.grapecity.com/spreadjs/hosted/css/gc.spread.sheets.excel2013white.14.2.6.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>

<body>
    <div class="page-wrapper">
        <div class="container">
            <header class="header">
                <h1 class="title">Werklijst Verkeersborden - SpreadJS</h1>
                <p class="description">
                    Enterprise-grade spreadsheet component met volledige Excel-compatibiliteit.
                </p>
                <div class="controls">
                    <a
                        href="https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1"
                        class="download-icon"
                        target="_blank"
                        rel="noopener noreferrer"
                        title="Bewerk het Excel-bestand"
                    ></a>
                    <button id="loadBtn" class="load-btn">Herlaad Data</button>
                    <button id="exportBtn" class="export-btn">Export Excel</button>
                    <button id="formatBtn" class="format-btn">Auto Format</button>
                    <button id="calcBtn" class="calc-btn">Herbereken</button>
                </div>
            </header>
            
            <div class="toolbar">
                <div class="toolbar-section">
                    <label>Zoom:</label>
                    <select id="zoomSelect">
                        <option value="0.5">50%</option>
                        <option value="0.75">75%</option>
                        <option value="1" selected>100%</option>
                        <option value="1.25">125%</option>
                        <option value="1.5">150%</option>
                        <option value="2">200%</option>
                    </select>
                </div>
                <div class="toolbar-section">
                    <label>Thema:</label>
                    <select id="themeSelect">
                        <option value="excel2013white">Excel 2013 White</option>
                        <option value="excel2013lightGray">Excel 2013 Light Gray</option>
                        <option value="excel2013darkGray">Excel 2013 Dark Gray</option>
                        <option value="excel2016colorful">Excel 2016 Colorful</option>
                    </select>
                </div>
                <div class="toolbar-section">
                    <span id="cellPosition">A1</span>
                    <span id="sheetInfo">Sheet1</span>
                </div>
            </div>
            
            <div id="loading" class="loading-indicator">
                Spreadsheet wordt geladen...
            </div>
            
            <div id="error" class="error-message" style="display: none;"></div>
            
            <div id="spreadjs" class="spreadjs-container"></div>
        </div>
    </div>

    <script>
        let spread = null;
        let workbook = null;
        let originalData = null;

        // Initialize SpreadJS
        function initializeSpreadJS() {
            try {
                // Create SpreadJS instance
                spread = new GC.Spread.Sheets.Workbook(document.getElementById('spreadjs'), {
                    sheetCount: 1,
                    allowUserResize: true,
                    allowUserZoom: true,
                    allowExtendPasteRange: true,
                    allowSheetReorder: true,
                    allowContextMenu: true,
                    allowUserEditFormula: true,
                    showHorizontalScrollbar: true,
                    showVerticalScrollbar: true,
                    showSheetTabs: true,
                    newTabVisible: false,
                    tabStripVisible: true,
                    tabNavigationVisible: true,
                    showResizeTip: true
                });

                workbook = spread;
                
                // Set up event handlers
                spread.bind(GC.Spread.Sheets.Events.CellClick, function (e, info) {
                    updateCellPosition(info.row, info.col);
                });

                spread.bind(GC.Spread.Sheets.Events.ActiveSheetChanged, function (e, info) {
                    updateSheetInfo();
                });

                spread.bind(GC.Spread.Sheets.Events.CellChanged, function (e, info) {
                    console.log('Cell changed:', info);
                });

                // Configure the active sheet
                const sheet = spread.getActiveSheet();
                sheet.name('Verkeersborden Data');
                
                // Set default row and column sizes
                sheet.defaults.rowHeight = 25;
                sheet.defaults.colWidth = 120;
                
                document.getElementById('loading').style.display = 'none';
                loadExcelData();
                
            } catch (error) {
                console.error('Error initializing SpreadJS:', error);
                showError('Fout bij het initialiseren van de spreadsheet: ' + error.message);
            }
        }

        function updateCellPosition(row, col) {
            const cellAddress = GC.Spread.Sheets.CalcEngine.rangeToFormula(new GC.Spread.Sheets.Range(row, col, 1, 1));
            document.getElementById('cellPosition').textContent = cellAddress;
        }

        function updateSheetInfo() {
            const sheet = spread.getActiveSheet();
            document.getElementById('sheetInfo').textContent = sheet.name();
        }

        async function loadExcelData() {
            try {
                const response = await fetch("https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1");
                if (!response.ok) {
                    throw new Error(`Het ophalen van het bestand is mislukt met status: ${response.status}`);
                }
                
                const arrayBuffer = await response.arrayBuffer();
                
                // Import Excel file directly into SpreadJS
                spread.import(arrayBuffer, function () {
                    console.log('Excel file loaded successfully');
                    
                    const sheet = spread.getActiveSheet();
                    
                    // Auto-fit columns
                    for (let i = 0; i < sheet.getColumnCount(); i++) {
                        sheet.autoFitColumn(i);
                    }
                    
                    // Apply header formatting
                    formatHeader();
                    
                    updateSheetInfo();
                    showSuccess('Excel bestand succesvol geladen!');
                    
                }, function (error) {
                    console.error('Import error:', error);
                    showError('Fout bij het importeren van Excel bestand: ' + error.message);
                });
                
            } catch (error) {
                console.error('Error loading Excel data:', error);
                showError('Kon de gegevens niet laden: ' + error.message);
            }
        }

        function formatHeader() {
            try {
                const sheet = spread.getActiveSheet();
                const usedRange = sheet.getUsedRange();
                
                if (usedRange && usedRange.rowCount > 0) {
                    // Format header row
                    const headerRange = new GC.Spread.Sheets.Range(0, 0, 1, usedRange.colCount);
                    const headerStyle = new GC.Spread.Sheets.Style();
                    headerStyle.backColor = '#4472C4';
                    headerStyle.foreColor = '#FFFFFF';
                    headerStyle.font = 'bold 12px Arial';
                    headerStyle.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
                    headerStyle.vAlign = GC.Spread.Sheets.VerticalAlign.center;
                    
                    sheet.setStyle(0, 0, headerStyle, GC.Spread.Sheets.SheetArea.viewport);
                    sheet.setRowHeight(0, 35, GC.Spread.Sheets.SheetArea.viewport);
                    
                    // Add borders
                    const borderStyle = new GC.Spread.Sheets.LineBorder('#000000', GC.Spread.Sheets.LineStyle.thin);
                    for (let col = 0; col < usedRange.colCount; col++) {
                        sheet.getCell(0, col, GC.Spread.Sheets.SheetArea.viewport).borderBottom(borderStyle);
                        sheet.getCell(0, col, GC.Spread.Sheets.SheetArea.viewport).borderTop(borderStyle);
                        sheet.getCell(0, col, GC.Spread.Sheets.SheetArea.viewport).borderLeft(borderStyle);
                        sheet.getCell(0, col, GC.Spread.Sheets.SheetArea.viewport).borderRight(borderStyle);
                    }
                }
            } catch (error) {
                console.error('Error formatting header:', error);
            }
        }

        function autoFormat() {
            try {
                const sheet = spread.getActiveSheet();
                const usedRange = sheet.getUsedRange();
                
                if (usedRange) {
                    // Auto-fit all columns
                    for (let i = 0; i < usedRange.colCount; i++) {
                        sheet.autoFitColumn(i);
                    }
                    
                    // Format data rows with alternating colors
                    for (let row = 1; row < usedRange.rowCount; row++) {
                        const rowStyle = new GC.Spread.Sheets.Style();
                        rowStyle.backColor = row % 2 === 0 ? '#F2F2F2' : '#FFFFFF';
                        
                        for (let col = 0; col < usedRange.colCount; col++) {
                            sheet.setStyle(row, col, rowStyle, GC.Spread.Sheets.SheetArea.viewport);
                        }
                    }
                    
                    formatHeader();
                    showSuccess('Auto-formattering toegepast!');
                }
            } catch (error) {
                showError('Fout bij auto-formattering: ' + error.message);
            }
        }

        function exportToExcel() {
            try {
                spread.export(function (blob) {
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'Verkeersborden_Export.xlsx';
                    a.click();
                    window.URL.revokeObjectURL(url);
                    showSuccess('Excel bestand wordt gedownload...');
                }, function (error) {
                    showError('Fout bij exporteren: ' + error.message);
                });
            } catch (error) {
                showError('Fout bij exporteren: ' + error.message);
            }
        }

        function recalculate() {
            try {
                const sheet = spread.getActiveSheet();
                sheet.recalcAll();
                showSuccess('Herberekening voltooid!');
            } catch (error) {
                showError('Fout bij herberekening: ' + error.message);
            }
        }

        function changeZoom() {
            const zoomLevel = parseFloat(document.getElementById('zoomSelect').value);
            spread.zoom(zoomLevel);
        }

        function changeTheme() {
            const theme = document.getElementById('themeSelect').value;
            // Note: Theme changing requires additional SpreadJS theme files
            console.log('Theme change requested:', theme);
            showSuccess('Thema wijziging niet beschikbaar in deze demo versie');
        }

        function showError(message) {
            const errorDiv = document.getElementById('error');
            errorDiv.textContent = message;
            errorDiv.style.display = 'block';
            setTimeout(() => {
                errorDiv.style.display = 'none';
            }, 5000);
        }

        function showSuccess(message) {
            const successDiv = document.createElement('div');
            successDiv.className = 'success-message';
            successDiv.textContent = message;
            successDiv.style.cssText = `
                position: fixed;
                top: 20px;
                right: 20px;
                background: #28a745;
                color: white;
                padding: 15px 20px;
                border-radius: 5px;
                z-index: 10000;
                box-shadow: 0 4px 8px rgba(0,0,0,0.2);
                animation: slideInRight 0.3s ease-out;
            `;
            document.body.appendChild(successDiv);
            
            setTimeout(() => {
                if (document.body.contains(successDiv)) {
                    document.body.removeChild(successDiv);
                }
            }, 3000);
        }

        // Event listeners
        document.addEventListener('DOMContentLoaded', () => {
            initializeSpreadJS();
            
            document.getElementById('loadBtn').addEventListener('click', loadExcelData);
            document.getElementById('exportBtn').addEventListener('click', exportToExcel);
            document.getElementById('formatBtn').addEventListener('click', autoFormat);
            document.getElementById('calcBtn').addEventListener('click', recalculate);
            document.getElementById('zoomSelect').addEventListener('change', changeZoom);
            document.getElementById('themeSelect').addEventListener('change', changeTheme);
        });

        // Handle window resize
        window.addEventListener('resize', () => {
            if (spread) {
                spread.refresh();
            }
        });
    </script>
</body>
</html>
