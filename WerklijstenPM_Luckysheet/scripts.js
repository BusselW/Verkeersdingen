/**
 * Luckysheet Excel Viewer - Modern Implementation
 * Enhanced with ES6+ features, async/await, and improved error handling
 */

class LuckysheetViewer {
    constructor(containerId) {
        this.containerId = containerId;
        this.isInitialized = false;
        this.currentData = null;
        this.fileName = null;
        
        this.init();
    }

    async init() {
        try {
            await this.waitForLibraries();
            this.setupUI();
            await this.initializeLuckysheet();
            await this.loadExcelData();
        } catch (error) {
            this.handleError(error);
        }
    }

    waitForLibraries() {
        return new Promise((resolve, reject) => {
            const checkLibraries = () => {
                if (typeof luckysheet !== 'undefined' && typeof XLSX !== 'undefined') {
                    resolve();
                } else {
                    setTimeout(checkLibraries, 100);
                }
            };
            checkLibraries();
            
            // Timeout after 10 seconds
            setTimeout(() => reject(new Error('Bibliotheken konden niet worden geladen')), 10000);
        });
    }

    setupUI() {
        const container = document.getElementById(this.containerId);
        if (!container) {
            throw new Error(`Container element '${this.containerId}' niet gevonden`);
        }

        container.innerHTML = `
            <div class="page-wrapper">
                <div class="container">
                    <header class="header">
                        <h1 class="title">Werklijst Verkeersborden - Luckysheet</h1>
                        <p class="description">
                            Interactieve spreadsheet viewer met volledige bewerkingsfunctionaliteit via Luckysheet.
                        </p>
                        <div class="controls">
                            <a
                                href="${this.excelUrl}"
                                class="download-icon"
                                target="_blank"
                                rel="noopener noreferrer"
                                title="Download origineel Excel bestand"
                            ></a>
                            <button id="loadBtn" class="action-btn">
                                <span class="btn-icon">🔄</span>
                                Herlaad Data
                            </button>
                            <button id="exportBtn" class="action-btn">
                                <span class="btn-icon">💾</span>
                                Export Excel
                            </button>
                        </div>
                    </header>
                    
                    <div id="notifications" class="notifications"></div>
                    <div id="spinner" class="spinner" style="display: none;">
                        <div class="spinner-content">
                            <div class="spinner-icon"></div>
                            <span>Bezig met laden...</span>
                        </div>
                    </div>
                    
                    <div class="spreadsheet-container">
                        <div id="luckysheet" style="margin:0px;padding:0px;position:absolute;width:100%;height:calc(100vh - 200px);left: 0px;top: 0px;"></div>
                    </div>
                </div>
            </div>
        `;

        // Attach event listeners
        document.getElementById('loadBtn').addEventListener('click', () => this.loadExcelData());
        document.getElementById('exportBtn').addEventListener('click', () => this.exportToExcel());
    }

    async initializeLuckysheet() {
        try {
            const options = {
                container: 'luckysheet',
                showinfobar: false,
                showsheetbar: true,
                showstatisticBar: true,
                data: [],
                title: 'Verkeersborden Werklijst',
                lang: 'en', // Using English as Dutch might not be fully supported
                allowCopy: true,
                allowEdit: true,
                allowUpdate: true
            };

            // Wait for DOM to be ready
            await new Promise(resolve => setTimeout(resolve, 100));
            
            luckysheet.create(options);
            
            // Wait for Luckysheet to be ready
            await new Promise((resolve) => {
                const checkReady = () => {
                    try {
                        if (luckysheet.getluckysheetfile && typeof luckysheet.getluckysheetfile() !== 'undefined') {
                            this.isInitialized = true;
                            this.showNotification('Luckysheet succesvol geïnitialiseerd!', 'success');
                            resolve();
                        } else {
                            setTimeout(checkReady, 100);
                        }
                    } catch (e) {
                        setTimeout(checkReady, 100);
                    }
                };
                checkReady();
            });
            
        } catch (error) {
            console.error('Error initializing Luckysheet:', error);
            throw new Error('Fout bij initialiseren van Luckysheet: ' + error.message);
        }
    }

    async loadExcelData() {
        try {
            this.showSpinner(true);
            this.showNotification('Excel bestand wordt geladen...', 'info');
            
            const response = await fetch(this.excelUrl);
            if (!response.ok) {
                throw new Error(`Het ophalen van het bestand is mislukt met status: ${response.status}`);
            }
            
            const arrayBuffer = await response.arrayBuffer();
            
            // Try Luckysheet's Excel import functionality if available
            if (typeof luckysheet.transformExcelToLucky === 'function') {
                await this.transformExcelData(arrayBuffer);
            } else {
                // Fallback to manual conversion
                await this.convertExcelManually(arrayBuffer);
            }
            
        } catch (error) {
            console.error('Error loading Excel data:', error);
            this.showNotification('Fout bij laden Excel bestand: ' + error.message, 'error');
        } finally {
            this.showSpinner(false);
        }
    }

    transformExcelData(arrayBuffer) {
        return new Promise((resolve, reject) => {
            try {
                luckysheet.transformExcelToLucky(arrayBuffer, (sheets, info) => {
                    try {
                        if (sheets && sheets.length > 0) {
                            // Clear existing data
                            try {
                                luckysheet.deleteSheet();
                            } catch (e) {
                                // Ignore delete errors if no sheets exist
                            }
                            
                            // Add new sheets
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
                            this.currentData = sheets;
                            this.showNotification(`Excel bestand succesvol geladen! ${sheets.length} werkblad(en)`, 'success');
                            resolve();
                        } else {
                            reject(new Error('Geen geldige spreadsheet data gevonden'));
                        }
                    } catch (err) {
                        reject(err);
                    }
                }, (error) => {
                    reject(new Error('Luckysheet transformatie fout: ' + error));
                });
            } catch (error) {
                reject(error);
            }
        });
    }

    async convertExcelManually(arrayBuffer) {
        try {
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { 
                type: "array",
                cellDates: true,
                cellNF: false,
                cellText: false
            });

            if (workbook.SheetNames.length === 0) {
                throw new Error("Het Excel-bestand bevat geen werkbladen.");
            }

            // Convert to Luckysheet format
            const sheets = workbook.SheetNames.map((sheetName, index) => {
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1,
                    defval: null,
                    blankrows: true
                });
                
                return this.convertToLuckysheetFormat(jsonData, sheetName, index);
            });

            // Load data into Luckysheet
            if (sheets.length > 0) {
                try {
                    luckysheet.deleteSheet();
                } catch (e) {
                    // Ignore delete errors if no sheets exist
                }
                
                sheets.forEach((sheet, index) => {
                    if (index === 0) {
                        luckysheet.setSheetData(sheet);
                    } else {
                        luckysheet.addSheet(sheet);
                    }
                });
                
                luckysheet.refresh();
                this.currentData = sheets;
                this.showNotification(`Excel bestand succesvol geladen! ${sheets.length} werkblad(en)`, 'success');
            } else {
                throw new Error('Geen data om te laden');
            }
            
        } catch (error) {
            throw new Error('Handmatige conversie fout: ' + error.message);
        }
    }

    convertToLuckysheetFormat(data, sheetName, index) {
        const celldata = [];
        
        data.forEach((row, rowIndex) => {
            if (Array.isArray(row)) {
                row.forEach((cell, colIndex) => {
                    if (cell !== null && cell !== undefined && cell !== '') {
                        celldata.push({
                            r: rowIndex,
                            c: colIndex,
                            v: {
                                v: cell,
                                ct: { fa: "General", t: "g" },
                                m: String(cell),
                                bg: rowIndex === 0 ? "#f8f9fa" : null,
                                fc: rowIndex === 0 ? "#212529" : null,
                                bl: rowIndex === 0 ? 1 : 0
                            }
                        });
                    }
                });
            }
        });

        return {
            name: sheetName || `Sheet${index + 1}`,
            color: '',
            index: index,
            status: index === 0 ? 1 : 0,
            order: index,
            hide: 0,
            row: Math.max(50, data.length + 10),
            column: Math.max(26, data[0] ? data[0].length + 5 : 26),
            defaultRowHeight: 19,
            defaultColWidth: 73,
            celldata: celldata,
            config: {},
            scrollLeft: 0,
            scrollTop: 0,
            luckysheet_select_save: [],
            calcChain: [],
            isPivotTable: false,
            pivotTable: {},
            filter_select: {},
            filter: null,
            luckysheet_alternateformat_save: [],
            luckysheet_alternateformat_save_modelCustom: [],
            luckysheet_conditionformat_save: {},
            frozen: {},
            chart: [],
            zoomRatio: 1,
            image: [],
            showGridLines: 1,
            dataVerification: {}
        };
    }

    exportToExcel() {
        try {
            if (!this.isInitialized) {
                this.showNotification('Luckysheet is nog niet geïnitialiseerd', 'error');
                return;
            }

            if (typeof luckysheet.exportLuckyToExcel === 'function') {
                luckysheet.exportLuckyToExcel('Verkeersborden_Export.xlsx');
                this.showNotification('Excel bestand wordt gedownload...', 'success');
            } else {
                // Fallback export method
                this.exportAsCSV();
            }
        } catch (error) {
            console.error('Export error:', error);
            this.showNotification('Fout bij exporteren: ' + error.message, 'error');
        }
    }

    exportAsCSV() {
        try {
            if (!this.currentData || this.currentData.length === 0) {
                this.showNotification('Geen data om te exporteren', 'error');
                return;
            }

            // Convert current sheet data to CSV
            const sheet = this.currentData[0];
            const rows = [];
            
            if (sheet && sheet.celldata) {
                const maxRow = Math.max(...sheet.celldata.map(cell => cell.r)) + 1;
                const maxCol = Math.max(...sheet.celldata.map(cell => cell.c)) + 1;
                
                for (let r = 0; r < maxRow; r++) {
                    const row = [];
                    for (let c = 0; c < maxCol; c++) {
                        const cell = sheet.celldata.find(item => item.r === r && item.c === c);
                        row.push(cell && cell.v && cell.v.v ? String(cell.v.v) : '');
                    }
                    rows.push(row);
                }
            }

            const csvData = rows.map(row => 
                row.map(cell => `"${String(cell || '').replace(/"/g, '""')}"`).join(',')
            ).join('\n');
            
            const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `verkeersborden_data_${new Date().toISOString().slice(0,10)}.csv`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            this.showNotification('CSV bestand wordt gedownload...', 'success');
        } catch (error) {
            this.showNotification('Fout bij CSV export: ' + error.message, 'error');
        }
    }

    showSpinner(show) {
        const spinner = document.getElementById('spinner');
        if (spinner) {
            spinner.style.display = show ? 'flex' : 'none';
        }
    }

    showNotification(message, type = 'info') {
        const container = document.getElementById('notifications');
        if (!container) return;

        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <span class="notification-message">${message}</span>
            <button class="notification-close">&times;</button>
        `;

        container.appendChild(notification);

        // Auto remove after 5 seconds
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 5000);

        // Manual close
        notification.querySelector('.notification-close').addEventListener('click', () => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        });
    }

    handleError(error) {
        console.error('LuckysheetViewer Error:', error);
        this.showNotification('Fout: ' + error.message, 'error');
        this.showSpinner(false);
    }
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    new LuckysheetViewer('root');
});

class LuckysheetViewer {
    constructor(containerId, excelUrl) {
        this.containerId = containerId;
        this.excelUrl = excelUrl;
        this.loading = true;
        this.error = null;
        this.luckysheetInstance = null;
        
        this.init();
    }

    async init() {
        try {
            await this.loadExcelData();
        } catch (error) {
            this.handleError(error);
        }
    }

    async loadExcelData() {
        try {
            this.setLoading(true);
            this.render();
            
            const response = await fetch(this.excelUrl);
            if (!response.ok) {
                throw new Error(`Het ophalen van het bestand is mislukt met status: ${response.status}`);
            }
            
            const arrayBuffer = await response.arrayBuffer();
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });
            
            // Convert all sheets to Luckysheet format
            const sheets = [];
            workbook.SheetNames.forEach((sheetName, index) => {
                const worksheet = workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1, 
                    defval: null,
                    raw: false
                });
                
                // Convert to Luckysheet format
                const luckysheetData = this.convertToLuckysheetFormat(sheetData, sheetName, index);
                sheets.push(luckysheetData);
            });
            
            if (sheets.length === 0) {
                throw new Error("Het Excel-bestand bevat geen geldige werkbladen.");
            }
            
            // Initialize Luckysheet
            this.initializeLuckysheet(sheets);
            
        } catch (error) {
            console.error("Fout bij het laden van Excel data:", error);
            this.handleError(error);
        }
    }

    convertToLuckysheetFormat(data, sheetName, index) {
        const celldata = [];
        
        data.forEach((row, rowIndex) => {
            if (Array.isArray(row)) {
                row.forEach((cell, colIndex) => {
                    if (cell !== null && cell !== undefined && cell !== '') {
                        celldata.push({
                            r: rowIndex,
                            c: colIndex,
                            v: {
                                v: cell,
                                ct: { fa: "General", t: "g" },
                                m: String(cell)
                            }
                        });
                    }
                });
            }
        });

        return {
            name: sheetName || `Werkblad ${index + 1}`,
            id: `sheet_${index}`,
            order: index,
            hide: 0,
            row: Math.max(data.length, 20),
            column: Math.max(data[0] ? data[0].length : 0, 10),
            defaultRowHeight: 19,
            defaultColWidth: 73,
            celldata: celldata,
            config: {
                merge: {},
                rowlen: {},
                columnlen: {},
                rowhidden: {},
                colhidden: {},
                borderInfo: [],
                authority: {}
            },
            scrollLeft: 0,
            scrollTop: 0,
            luckysheet_select_save: [],
            calcChain: [],
            isPivotTable: false,
            pivotTable: {},
            filter_select: {},
            filter: null,
            luckysheet_alternateformat_save: [],
            luckysheet_alternateformat_save_modelCustom: [],
            luckysheet_conditionformat_save: {},
            frozen: {},
            chart: [],
            zoomRatio: 1,
            image: [],
            showGridLines: 1,
            dataVerification: {}
        };
    }

    initializeLuckysheet(sheets) {
        try {
            const options = {
                container: 'luckysheet',
                data: sheets,
                title: 'Verkeersborden Werklijst',
                lang: 'nl',
                allowCopy: true,
                allowEdit: false,
                allowDelete: false,
                allowAdd: false,
                showToolbar: true,
                showInfoBar: true,
                showSheetBarConfig: {
                    add: false,
                    menu: false,
                    sheet: true
                },
                enableAddRow: false,
                enableAddCol: false,
                sheetRightClickConfig: {
                    delete: false,
                    copy: false,
                    rename: false,
                    color: false,
                    hide: false,
                    move: false
                },
                cellRightClickConfig: {
                    copy: true,
                    copyAs: false,
                    paste: false,
                    insertRow: false,
                    insertColumn: false,
                    deleteRow: false,
                    deleteColumn: false,
                    deleteCell: false,
                    hideRow: false,
                    hideColumn: false,
                    rowHeight: false,
                    columnWidth: false,
                    clear: false,
                    matrix: false,
                    sort: false,
                    filter: false,
                    chart: false,
                    image: false,
                    link: false,
                    data: false,
                    cellFormat: false
                },
                hook: {
                    workbookCreateAfter: () => {
                        this.setLoading(false);
                        this.showNotification('Excel bestand succesvol geladen!', 'success');
                    },
                    workbookCreateBefore: () => {
                        console.log('Luckysheet wordt geïnitialiseerd...');
                    }
                }
            };

            // Clear any existing Luckysheet instance
            if (window.luckysheet && typeof window.luckysheet.destroy === 'function') {
                window.luckysheet.destroy();
            }

            // Initialize Luckysheet
            window.luckysheet.create(options);
            this.luckysheetInstance = window.luckysheet;
            
        } catch (error) {
            console.error("Fout bij het initialiseren van Luckysheet:", error);
            this.handleError(new Error("Kon de spreadsheet viewer niet initialiseren."));
        }
    }

    setLoading(loading) {
        this.loading = loading;
        if (loading) {
            this.render();
        }
    }

    handleError(error) {
        console.error("Luckysheet Viewer Error:", error);
        this.error = error.message;
        this.loading = false;
        this.render();
    }

    exportToExcel() {
        try {
            if (!this.luckysheetInstance) {
                throw new Error("Spreadsheet niet beschikbaar voor export");
            }
            
            // Use Luckysheet's built-in export functionality
            this.luckysheetInstance.exportXlsx('verkeersborden_export.xlsx');
            this.showNotification('Excel export gestart!', 'success');
        } catch (error) {
            this.showNotification('Fout bij Excel export: ' + error.message, 'error');
        }
    }

    exportToCSV() {
        try {
            if (!this.luckysheetInstance) {
                throw new Error("Spreadsheet niet beschikbaar voor export");
            }
            
            const sheetData = this.luckysheetInstance.getSheetData();
            const csvData = sheetData.map(row => 
                row.map(cell => {
                    const value = cell && cell.v ? cell.v : '';
                    return `"${String(value).replace(/"/g, '""')}"`;
                }).join(',')
            ).join('\n');
            
            const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `verkeersborden_export_${new Date().toISOString().slice(0,10)}.csv`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            this.showNotification('CSV export succesvol!', 'success');
        } catch (error) {
            this.showNotification('Fout bij CSV export: ' + error.message, 'error');
        }
    }

    showNotification(message, type = 'info') {
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.textContent = message;
        
        Object.assign(notification.style, {
            position: 'fixed',
            top: '20px',
            right: '20px',
            padding: '12px 20px',
            borderRadius: '8px',
            color: 'white',
            fontWeight: '500',
            zIndex: '10000',
            transform: 'translateX(400px)',
            transition: 'transform 0.3s ease',
            backgroundColor: type === 'success' ? '#28a745' : type === 'error' ? '#dc3545' : '#17a2b8'
        });
        
        document.body.appendChild(notification);
        
        setTimeout(() => {
            notification.style.transform = 'translateX(0)';
        }, 100);
        
        setTimeout(() => {
            notification.style.transform = 'translateX(400px)';
            setTimeout(() => {
                if (document.body.contains(notification)) {
                    document.body.removeChild(notification);
                }
            }, 300);
        }, 3000);
    }

    render() {
        const container = document.getElementById(this.containerId);
        if (!container) return;

        if (this.loading) {
            container.innerHTML = `
                <div class="page-wrapper">
                    <div class="container">
                        <header class="header">
                            <h1 class="title">Werklijst Verkeersborden - Luckysheet</h1>
                            <p class="description">
                                Geavanceerde spreadsheet viewer met volledige Excel functionaliteit.
                            </p>
                        </header>
                        <div class="loading-indicator">
                            <div class="spinner"></div>
                            <p>Excel bestand laden...</p>
                        </div>
                    </div>
                </div>
            `;
            return;
        }

        if (this.error) {
            container.innerHTML = `
                <div class="page-wrapper">
                    <div class="container">
                        <header class="header">
                            <h1 class="title">Werklijst Verkeersborden - Luckysheet</h1>
                            <p class="description">
                                Geavanceerde spreadsheet viewer met volledige Excel functionaliteit.
                            </p>
                        </header>
                        <div class="error-message">${this.escapeHtml(this.error)}</div>
                        <div style="text-align: center; padding: 20px;">
                            <button onclick="location.reload()" class="load-btn">Opnieuw proberen</button>
                        </div>
                    </div>
                </div>
            `;
            return;
        }

        container.innerHTML = `
            <div class="page-wrapper">
                <div class="container">
                    <header class="header">
                        <h1 class="title">Werklijst Verkeersborden - Luckysheet</h1>
                        <p class="description">
                            Geavanceerde spreadsheet viewer met volledige Excel functionaliteit en interactieve editing.
                        </p>
                        <div class="controls">
                            <a
                                href="${this.excelUrl}"
                                class="download-icon"
                                target="_blank"
                                rel="noopener noreferrer"
                                title="Download origineel Excel bestand"
                                aria-label="Download Excel bestand"
                            ></a>
                            <button onclick="luckysheetViewer.exportToCSV()" class="export-btn">
                                📄 Export CSV
                            </button>
                            <button onclick="luckysheetViewer.exportToExcel()" class="export-btn">
                                📊 Export Excel
                            </button>
                        </div>
                    </header>
                    <div class="luckysheet-container">
                        <div id="luckysheet" style="margin:0px;padding:0px;position:absolute;width:100%;height:100%;left: 0px;top: 0px;"></div>
                    </div>
                </div>
            </div>
        `;
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    destroy() {
        if (this.luckysheetInstance && typeof this.luckysheetInstance.destroy === 'function') {
            this.luckysheetInstance.destroy();
        }
    }
}

// Initialize when DOM is ready and Luckysheet is loaded
function initializeLuckysheetViewer() {
    if (typeof window.luckysheet !== 'undefined') {
        const excelUrl = "https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1";
        window.luckysheetViewer = new LuckysheetViewer('root', excelUrl);
    } else {
        // Retry after a short delay if Luckysheet isn't loaded yet
        setTimeout(initializeLuckysheetViewer, 100);
    }
}

document.addEventListener('DOMContentLoaded', initializeLuckysheetViewer);
