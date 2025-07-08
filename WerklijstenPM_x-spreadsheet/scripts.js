/**
 * x-spreadsheet Excel Viewer - Modern Implementation
 * Enhanced with ES6+ features, async/await, and improved error handling
 */

class XSpreadsheetViewer {
    constructor(containerId, excelUrl) {
        this.containerId = containerId;
        this.excelUrl = excelUrl;
        this.xs = null;
        this.originalData = null;
        this.isInitialized = false;
        
        this.init();
    }

    async init() {
        try {
            await this.waitForLibraries();
            this.setupUI();
            this.initializeSpreadsheet();
            await this.loadExcelData();
        } catch (error) {
            this.handleError(error);
        }
    }

    waitForLibraries() {
        return new Promise((resolve, reject) => {
            const checkLibraries = () => {
                if (typeof x_spreadsheet !== 'undefined' && typeof XLSX !== 'undefined') {
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
                        <h1 class="title">Werklijst Verkeersborden - x-spreadsheet</h1>
                        <p class="description">
                            Lichtgewicht, snel en moderne spreadsheet viewer met x-spreadsheet.
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
                                Export CSV
                            </button>
                            <button id="clearBtn" class="action-btn secondary">
                                <span class="btn-icon">🗑️</span>
                                Wis Sheet
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
                        <div id="spreadsheet" style="width: 100%; height: calc(100vh - 200px);"></div>
                    </div>
                </div>
            </div>
        `;

        // Attach event listeners
        document.getElementById('loadBtn').addEventListener('click', () => this.loadExcelData());
        document.getElementById('exportBtn').addEventListener('click', () => this.exportData());
        document.getElementById('clearBtn').addEventListener('click', () => this.clearSheet());
    }

    initializeSpreadsheet() {
        try {
            const options = {
                mode: 'edit',
                showToolbar: true,
                showGrid: true,
                showContextmenu: true,
                view: {
                    height: () => {
                        const container = document.getElementById('spreadsheet');
                        return container ? container.clientHeight : 600;
                    },
                    width: () => {
                        const container = document.getElementById('spreadsheet');
                        return container ? container.clientWidth : 800;
                    },
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
                },
            };

            this.xs = x_spreadsheet('#spreadsheet', options);
            this.isInitialized = true;
            this.showNotification('x-spreadsheet succesvol geïnitialiseerd!', 'success');
            
            // Handle window resize properly
            this.setupResizeHandler();
            
        } catch (error) {
            console.error('Error initializing x-spreadsheet:', error);
            throw new Error('Fout bij initialiseren spreadsheet: ' + error.message);
        }
    }

    setupResizeHandler() {
        let resizeTimeout;
        window.addEventListener('resize', () => {
            clearTimeout(resizeTimeout);
            resizeTimeout = setTimeout(() => {
                try {
                    if (this.xs && typeof this.xs.resize === 'function') {
                        this.xs.resize();
                    } else if (this.xs && this.xs.reload) {
                        // Alternative method if resize doesn't exist
                        this.xs.reload();
                    }
                } catch (error) {
                    console.warn('Resize failed:', error);
                    // Silently ignore resize errors as they're not critical
                }
            }, 250);
        });
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

            // Process the first sheet
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,
                defval: '',
                blankrows: false
            });
            
            if (jsonData.length === 0) {
                throw new Error("Het Excel-bestand is leeg of kon niet correct worden gelezen.");
            }

            // Store original data
            this.originalData = jsonData;

            // Convert to x-spreadsheet format
            const spreadsheetData = this.convertToXSpreadsheetFormat(jsonData);
            
            if (this.xs && this.isInitialized) {
                this.xs.loadData(spreadsheetData);
                this.showNotification(`Excel bestand succesvol geladen! ${jsonData.length} rijen, ${jsonData[0] ? jsonData[0].length : 0} kolommen`, 'success');
            } else {
                throw new Error('x-spreadsheet is niet geïnitialiseerd');
            }
            
        } catch (error) {
            console.error('Error loading Excel data:', error);
            this.showNotification('Kon de gegevens niet laden: ' + error.message, 'error');
        } finally {
            this.showSpinner(false);
        }
    }

    convertToXSpreadsheetFormat(data) {
        const rows = {};
        
        data.forEach((row, rowIndex) => {
            const cells = {};
            row.forEach((cell, colIndex) => {
                if (cell !== null && cell !== undefined && cell !== '') {
                    cells[colIndex] = {
                        text: String(cell),
                        style: rowIndex === 0 ? 1 : 0 // Header style
                    };
                }
            });
            if (Object.keys(cells).length > 0) {
                rows[rowIndex] = { cells };
            }
        });

        return {
            name: 'Verkeersborden',
            freeze: 'A1',
            styles: [
                {
                    bgcolor: '#f8f9fa',
                    color: '#212529',
                    bold: true,
                    align: 'center'
                },
                {
                    bgcolor: '#ffffff',
                    color: '#212529',
                    align: 'left'
                }
            ],
            merges: [],
            cols: data[0] ? data[0].reduce((acc, _, index) => {
                acc[index] = { width: 120 };
                return acc;
            }, {}) : {},
            rows: rows,
            validations: []
        };
    }

    exportData() {
        try {
            if (!this.originalData || this.originalData.length === 0) {
                this.showNotification('Geen data om te exporteren. Laad eerst een Excel-bestand.', 'error');
                return;
            }

            // Export as CSV
            const csvData = this.originalData.map(row => 
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
            console.error('Export error:', error);
            this.showNotification('Fout bij exporteren: ' + error.message, 'error');
        }
    }

    clearSheet() {
        try {
            if (this.xs && this.isInitialized) {
                this.xs.loadData({
                    name: 'Verkeersborden',
                    freeze: 'A1',
                    styles: [],
                    merges: [],
                    cols: {},
                    rows: {},
                    validations: []
                });
                this.originalData = null;
                this.showNotification('Spreadsheet geleegd!', 'success');
            } else {
                this.showNotification('x-spreadsheet is niet geïnitialiseerd', 'error');
            }
        } catch (error) {
            console.error('Clear error:', error);
            this.showNotification('Fout bij legen spreadsheet: ' + error.message, 'error');
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
        console.error('XSpreadsheetViewer Error:', error);
        this.showNotification('Fout: ' + error.message, 'error');
        this.showSpinner(false);
    }
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    const excelUrl = "https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1";
    new XSpreadsheetViewer('root', excelUrl);
});
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
            
            if (workbook.SheetNames.length === 0) {
                throw new Error("Het Excel-bestand bevat geen werkbladen.");
            }
            
            // Convert to x-spreadsheet format
            const spreadsheetData = this.convertToXSpreadsheetFormat(workbook);
            this.currentData = spreadsheetData;
            
            // Initialize x-spreadsheet
            this.initializeSpreadsheet(spreadsheetData);
            
            this.setLoading(false);
            this.showNotification('Excel bestand succesvol geladen!', 'success');
            
        } catch (error) {
            console.error("Fout bij het laden van Excel data:", error);
            this.handleError(error);
        }
    }

    convertToXSpreadsheetFormat(workbook) {
        const sheets = [];
        
        workbook.SheetNames.forEach((sheetName, index) => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1, 
                defval: '',
                raw: false 
            });
            
            // Convert to x-spreadsheet format
            const rows = {};
            jsonData.forEach((row, rowIndex) => {
                if (Array.isArray(row)) {
                    const cells = {};
                    row.forEach((cell, colIndex) => {
                        if (cell !== null && cell !== undefined && cell !== '') {
                            cells[colIndex] = { text: String(cell) };
                        }
                    });
                    if (Object.keys(cells).length > 0) {
                        rows[rowIndex] = { cells };
                    }
                }
            });
            
            sheets.push({
                name: sheetName || `Werkblad ${index + 1}`,
                rows: rows,
                merges: [],
                styles: [],
                validations: [],
                cols: {},
                freeze: 'A1'
            });
        });
        
        return sheets;
    }

    initializeSpreadsheet(data) {
        try {
            const container = document.getElementById('spreadsheet');
            if (!container) {
                throw new Error("Spreadsheet container niet gevonden");
            }
            
            // Clear any existing content
            container.innerHTML = '';
            
            // Configure x-spreadsheet options
            const options = {
                mode: 'read', // Set to read-only by default
                showToolbar: true,
                showGrid: true,
                showContextmenu: false,
                view: {
                    height: () => container.offsetHeight - 40,
                    width: () => container.offsetWidth - 20
                },
                row: {
                    len: 1000,
                    height: 25
                },
                col: {
                    len: 26,
                    width: 100,
                    indexWidth: 60,
                    minWidth: 60
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
                        name: 'Inter, Arial',
                        size: 10,
                        bold: false,
                        italic: false
                    }
                }
            };
            
            // Initialize x-spreadsheet
            this.spreadsheetInstance = x_spreadsheet(container, options);
            
            // Load data
            if (data && data.length > 0) {
                this.spreadsheetInstance.loadData(data);
            }
            
            // Handle resize
            this.setupResize();
            
        } catch (error) {
            console.error("Fout bij het initialiseren van x-spreadsheet:", error);
            this.handleError(new Error("Kon de spreadsheet viewer niet initialiseren."));
        }
    }

    setupResize() {
        const resizeHandler = () => {
            if (this.spreadsheetInstance && typeof this.spreadsheetInstance.resize === 'function') {
                try {
                    this.spreadsheetInstance.resize();
                } catch (error) {
                    console.warn("Resize warning:", error.message);
                }
            }
        };
        
        window.addEventListener('resize', resizeHandler);
        
        // Store cleanup function
        this.cleanupResize = () => {
            window.removeEventListener('resize', resizeHandler);
        };
    }

    setLoading(loading) {
        this.loading = loading;
        if (loading) {
            this.render();
        }
    }

    handleError(error) {
        console.error("x-spreadsheet Viewer Error:", error);
        this.error = error.message;
        this.loading = false;
        this.render();
    }

    exportToCSV() {
        try {
            if (!this.currentData || this.currentData.length === 0) {
                throw new Error("Geen data beschikbaar voor export");
            }
            
            // Get current sheet data
            const sheetData = this.currentData[0]; // Use first sheet
            const csvRows = [];
            
            // Convert rows object to array format
            const maxRow = Math.max(...Object.keys(sheetData.rows).map(Number));
            for (let i = 0; i <= maxRow; i++) {
                const row = sheetData.rows[i];
                const csvRow = [];
                
                if (row && row.cells) {
                    const maxCol = Math.max(...Object.keys(row.cells).map(Number));
                    for (let j = 0; j <= maxCol; j++) {
                        const cell = row.cells[j];
                        const value = cell ? cell.text || '' : '';
                        csvRow.push(`"${String(value).replace(/"/g, '""')}"`);
                    }
                }
                csvRows.push(csvRow.join(','));
            }
            
            const csvData = csvRows.join('\n');
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

    exportToJSON() {
        try {
            if (!this.currentData || this.currentData.length === 0) {
                throw new Error("Geen data beschikbaar voor export");
            }
            
            const jsonData = this.currentData.map(sheet => ({
                name: sheet.name,
                data: sheet.rows
            }));
            
            const blob = new Blob([JSON.stringify(jsonData, null, 2)], { type: 'application/json' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `verkeersborden_export_${new Date().toISOString().slice(0,10)}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            this.showNotification('JSON export succesvol!', 'success');
        } catch (error) {
            this.showNotification('Fout bij JSON export: ' + error.message, 'error');
        }
    }

    clearSpreadsheet() {
        try {
            if (this.spreadsheetInstance) {
                this.spreadsheetInstance.loadData([{
                    name: 'Sheet1',
                    rows: {},
                    merges: [],
                    styles: [],
                    validations: [],
                    cols: {}
                }]);
                this.showNotification('Spreadsheet gewist!', 'success');
            }
        } catch (error) {
            this.showNotification('Fout bij wissen: ' + error.message, 'error');
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
            backgroundColor: type === 'success' ? '#28a745' : type === 'error' ? '#dc3545' : '#6f42c1'
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
                            <h1 class="title">Werklijst Verkeersborden - x-spreadsheet</h1>
                            <p class="description">
                                Lichtgewicht en snelle spreadsheet viewer met moderne functionaliteit.
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
                            <h1 class="title">Werklijst Verkeersborden - x-spreadsheet</h1>
                            <p class="description">
                                Lichtgewicht en snelle spreadsheet viewer met moderne functionaliteit.
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
                        <h1 class="title">Werklijst Verkeersborden - x-spreadsheet</h1>
                        <p class="description">
                            Lichtgewicht en snelle spreadsheet viewer met moderne functionaliteit en export opties.
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
                            <button onclick="xSpreadsheetViewer.exportToCSV()" class="export-btn">
                                📄 Export CSV
                            </button>
                            <button onclick="xSpreadsheetViewer.exportToJSON()" class="export-btn">
                                📋 Export JSON
                            </button>
                            <button onclick="xSpreadsheetViewer.clearSpreadsheet()" class="clear-btn">
                                🗑️ Wissen
                            </button>
                        </div>
                    </header>
                    <div class="info-bar">
                        <span>📊 Status: Geladen</span>
                        <span>🔧 Type: x-spreadsheet</span>
                        <span>📋 Werkbladen: ${this.currentData ? this.currentData.length : 0}</span>
                    </div>
                    <div class="spreadsheet-container">
                        <div id="spreadsheet"></div>
                    </div>
                </div>
            </div>
        `;
        
        // Initialize spreadsheet after render
        setTimeout(() => {
            if (this.currentData) {
                this.initializeSpreadsheet(this.currentData);
            }
        }, 100);
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    destroy() {
        if (this.cleanupResize) {
            this.cleanupResize();
        }
        if (this.spreadsheetInstance) {
            this.spreadsheetInstance = null;
        }
    }
}

// Initialize when DOM is ready and x-spreadsheet is loaded
function initializeXSpreadsheetViewer() {
    if (typeof window.x_spreadsheet !== 'undefined') {
        const excelUrl = "https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1";
        window.xSpreadsheetViewer = new XSpreadsheetViewer('root', excelUrl);
    } else {
        // Retry after a short delay if x-spreadsheet isn't loaded yet
        setTimeout(initializeXSpreadsheetViewer, 100);
    }
}

document.addEventListener('DOMContentLoaded', initializeXSpreadsheetViewer);
