<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verkeersborden Werklijst - Luckysheet</title>
    <link href="styles.css" rel="stylesheet">
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/css/pluginsCss.css' />
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/plugins.css' />
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/css/luckysheet.css' />
    <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/assets/iconfont/iconfont.css' />
    <script src="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/js/plugin.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/luckysheet.umd.js"></script>
</head>

<body>
    <div class="page-wrapper">
        <div class="container">
            <header class="header">
                <h1 class="title">Werklijst Verkeersborden - Luckysheet</h1>
                <p class="description">
                    Volledig interactieve spreadsheet editor met Excel-compatibiliteit.
                </p>
                <div class="controls">
                    <a
                        href="https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1"
                        class="download-icon"
                        target="_blank"
                        rel="noopener noreferrer"
                        title="Bewerk het Excel-bestand"
                    ></a>
                    <button id="loadBtn" class="load-btn">Laad Excel Bestand</button>
                    <button id="exportBtn" class="export-btn">Export Excel</button>
                </div>
            </header>
            
            <div id="loading" class="loading-indicator">
                Spreadsheet wordt geladen...
            </div>
            
            <div id="error" class="error-message" style="display: none;"></div>
            
            <div id="luckysheet" class="luckysheet-container"></div>
        </div>
    </div>

    <script>
        let workbookData = [];
        
        // Initialize Luckysheet
        const options = {
            container: 'luckysheet',
            title: 'Verkeersborden Werklijst',
            lang: 'nl',
            allowCopy: true,
            allowEdit: true,
            allowDelete: true,
            showsheetbar: true,
            showstatisticBar: true,
            enableAddRow: true,
            enableAddCol: true,
            userInfo: false,
            myFolderUrl: '',
            loadUrl: '',
            updateUrl: '',
            loadSheetUrl: '',
            allowUpdate: false,
            functionButton: '',
            showConfigWindowResize: false,
            forceCalculation: false,
            data: [],
            hook: {
                workbookCreateAfter: function() {
                    document.getElementById('loading').style.display = 'none';
                    loadExcelData();
                },
                cellEditBefore: function(range) {
                    // Allow editing
                    return true;
                },
                sheetActivate: function(index, isPivotInitial, isNewSheet) {
                    console.log('Sheet activated:', index);
                }
            }
        };

        // Initialize Luckysheet
        luckysheet.create(options);

        async function loadExcelData() {
            try {
                const response = await fetch("https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1");
                if (!response.ok) {
                    throw new Error(`Het ophalen van het bestand is mislukt met status: ${response.status}`);
                }
                
                const arrayBuffer = await response.arrayBuffer();
                
                // Convert Excel file to Luckysheet format
                luckysheet.transformExcelToLucky(arrayBuffer, function(sheets, info) {
                    if (sheets && sheets.length > 0) {
                        // Clear existing sheets and load new ones
                        luckysheet.deleteSheet();
                        
                        sheets.forEach((sheet, index) => {
                            if (index === 0) {
                                // Replace the first sheet
                                luckysheet.setSheetData(sheet);
                            } else {
                                // Add additional sheets
                                luckysheet.addSheet(sheet);
                            }
                        });
                        
                        luckysheet.refresh();
                        showSuccess('Excel bestand succesvol geladen!');
                    } else {
                        throw new Error('Geen geldige spreadsheet data gevonden');
                    }
                }, function(error) {
                    showError('Fout bij het laden van Excel bestand: ' + error);
                });
                
            } catch (error) {
                console.error('Error loading Excel data:', error);
                showError('Kon de gegevens niet laden: ' + error.message);
            }
        }

        function exportToExcel() {
            try {
                const sheets = luckysheet.getAllSheets();
                if (sheets && sheets.length > 0) {
                    luckysheet.exportLuckyToExcel('Verkeersborden_Export.xlsx');
                    showSuccess('Excel bestand wordt gedownload...');
                } else {
                    showError('Geen data om te exporteren');
                }
            } catch (error) {
                console.error('Export error:', error);
                showError('Fout bij exporteren: ' + error.message);
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
            // Create temporary success message
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
            `;
            document.body.appendChild(successDiv);
            
            setTimeout(() => {
                document.body.removeChild(successDiv);
            }, 3000);
        }

        // Event listeners
        document.getElementById('loadBtn').addEventListener('click', loadExcelData);
        document.getElementById('exportBtn').addEventListener('click', exportToExcel);
    </script>
</body>
</html>
