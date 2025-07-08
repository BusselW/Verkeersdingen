<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verkeersborden Werklijst - x-spreadsheet</title>
    <link href="styles.css" rel="stylesheet">
    <script src="https://unpkg.com/x-data-spreadsheet@1.1.9/dist/xspreadsheet.js"></script>
    <link rel="stylesheet" href="https://unpkg.com/x-data-spreadsheet@1.1.9/dist/xspreadsheet.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>

<body>
    <div class="page-wrapper">
        <div class="container">
            <header class="header">
                <h1 class="title">Werklijst Verkeersborden - x-spreadsheet</h1>
                <p class="description">
                    Moderne, lightweight spreadsheet component met canvas rendering.
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
                    <button id="exportBtn" class="export-btn">Export JSON</button>
                    <button id="clearBtn" class="clear-btn">Leeg Sheet</button>
                </div>
            </header>
            
            <div class="info-bar">
                <span id="cellInfo">Cel: A1</span>
                <span id="sheetInfo">Sheet: 1</span>
                <span id="dataInfo">Rijen: 0, Kolommen: 0</span>
            </div>
            
            <div id="loading" class="loading-indicator">
                Spreadsheet wordt geladen...
            </div>
            
            <div id="error" class="error-message" style="display: none;"></div>
            
            <div id="spreadsheet" class="spreadsheet-container"></div>
        </div>
    </div>

    <script>
        let xs = null;
        let originalData = null;

        // Initialize x-spreadsheet
        function initializeSpreadsheet() {
            const options = {
                mode: 'edit',
                showToolbar: true,
                showGrid: true,
                showContextmenu: true,
                view: {
                    height: () => document.getElementById('spreadsheet').clientHeight,
                    width: () => document.getElementById('spreadsheet').clientWidth,
                },
                row: {
                    len: 100,
                    height: 25,
                },
                col: {
                    len: 26,
                    width: 100,
                    indexWidth: 60,
                    minWidth: 60,
                },
                style: {
                    bgcolor: '#ffffff',
                    align: 'left',
                    valign: 'middle',
                    textwrap: false,
                    strike: false,
                    underline: false,
                    color: '#0a0a0a',
                    font: {
                        name: 'Helvetica',
                        size: 10,
                        bold: false,
                        italic: false,
                    },
                },
            };

            xs = x_spreadsheet('#spreadsheet', options);
            
            // Add event listeners
            xs.on('cell-selected', (cell, ri, ci) => {
                updateCellInfo(ri, ci);
            });
            
            xs.on('cell-edited', (text, ri, ci) => {
                console.log('Cell edited:', text, ri, ci);
                updateDataInfo();
            });

            document.getElementById('loading').style.display = 'none';
            loadExcelData();
        }

        function updateCellInfo(ri, ci) {
            const cellName = columnName(ci) + (ri + 1);
            document.getElementById('cellInfo').textContent = `Cel: ${cellName}`;
        }

        function updateDataInfo() {
            try {
                const data = xs.getData();
                const sheets = Object.keys(data);
                const currentSheet = sheets[0] || 'Sheet1';
                const sheetData = data[currentSheet];
                
                let maxRow = 0;
                let maxCol = 0;
                
                if (sheetData && sheetData.rows) {
                    Object.keys(sheetData.rows).forEach(key => {
                        const rowIndex = parseInt(key);
                        if (rowIndex > maxRow) maxRow = rowIndex;
                        
                        const row = sheetData.rows[key];
                        if (row && row.cells) {
                            Object.keys(row.cells).forEach(cellKey => {
                                const colIndex = parseInt(cellKey);
                                if (colIndex > maxCol) maxCol = colIndex;
                            });
                        }
                    });
                }
                
                document.getElementById('dataInfo').textContent = `Rijen: ${maxRow + 1}, Kolommen: ${maxCol + 1}`;
                document.getElementById('sheetInfo').textContent = `Sheet: ${currentSheet}`;
            } catch (error) {
                console.error('Error updating data info:', error);
            }
        }

        function columnName(index) {
            let result = '';
            while (index >= 0) {
                result = String.fromCharCode(65 + (index % 26)) + result;
                index = Math.floor(index / 26) - 1;
            }
            return result;
        }

        async function loadExcelData() {
            try {
                const response = await fetch("https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1");
                if (!response.ok) {
                    throw new Error(`Het ophalen van het bestand is mislukt met status: ${response.status}`);
                }
                
                const arrayBuffer = await response.arrayBuffer();
                const data = new Uint8Array(arrayBuffer);
                const workbook = XLSX.read(data, { type: "array" });
                
                // Convert XLSX data to x-spreadsheet format
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                
                if (jsonData.length === 0) {
                    throw new Error("Het Excel-bestand is leeg of kon niet correct worden gelezen.");
                }

                // Convert to x-spreadsheet format
                const xsData = convertToXSpreadsheetFormat(jsonData);
                originalData = jsonData;
                
                // Load data into spreadsheet
                xs.loadData(xsData);
                updateDataInfo();
                showSuccess('Excel bestand succesvol geladen!');
                
            } catch (error) {
                console.error('Error loading Excel data:', error);
                showError('Kon de gegevens niet laden: ' + error.message);
            }
        }

        function convertToXSpreadsheetFormat(jsonData) {
            const rows = {};
            
            jsonData.forEach((row, rowIndex) => {
                if (row && row.length > 0) {
                    const cells = {};
                    row.forEach((cell, colIndex) => {
                        if (cell !== null && cell !== undefined && cell !== '') {
                            cells[colIndex] = {
                                text: String(cell),
                                style: 0
                            };
                        }
                    });
                    
                    if (Object.keys(cells).length > 0) {
                        rows[rowIndex] = { cells };
                    }
                }
            });

            return {
                'Sheet1': {
                    name: 'Verkeersborden',
                    freeze: 'A1',
                    styles: [
                        {
                            align: 'center',
                            valign: 'middle',
                            font: { bold: true },
                            bgcolor: '#f0f0f0'
                        }
                    ],
                    merges: [],
                    rows: rows,
                    cols: {}
                }
            };
        }

        function exportData() {
            try {
                const data = xs.getData();
                const jsonString = JSON.stringify(data, null, 2);
                const blob = new Blob([jsonString], { type: 'application/json' });
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'verkeersborden_data.json';
                a.click();
                window.URL.revokeObjectURL(url);
                showSuccess('Data geëxporteerd als JSON!');
            } catch (error) {
                showError('Fout bij exporteren: ' + error.message);
            }
        }

        function clearSheet() {
            if (confirm('Weet je zeker dat je de spreadsheet wilt legen?')) {
                xs.loadData({
                    'Sheet1': {
                        name: 'Sheet1',
                        rows: {},
                        cols: {}
                    }
                });
                updateDataInfo();
                showSuccess('Spreadsheet geleegd!');
            }
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
            initializeSpreadsheet();
            
            document.getElementById('loadBtn').addEventListener('click', loadExcelData);
            document.getElementById('exportBtn').addEventListener('click', exportData);
            document.getElementById('clearBtn').addEventListener('click', clearSheet);
        });

        // Handle window resize
        window.addEventListener('resize', () => {
            if (xs) {
                xs.resize();
            }
        });
    </script>
</body>
</html>
