/**
 * SheetJS Excel Viewer - Modern Implementation
 * Enhanced with ES6+ features, async/await, and improved error handling
 */

class ExcelViewer {
    constructor(containerId, excelUrl) {
        this.containerId = containerId;
        this.excelUrl = excelUrl;
        this.tableData = [];
        this.sheetNames = [];
        this.currentSheet = 0;
        this.loading = true;
        this.error = null;
        
        // Bind methods
        this.handleSheetChange = this.handleSheetChange.bind(this);
        this.exportToCSV = this.exportToCSV.bind(this);
        this.exportToJSON = this.exportToJSON.bind(this);
        
        this.init();
    }

    async init() {
        try {
            await this.loadExcelData();
            this.setupScrollSynchronization();
            this.render();
        } catch (error) {
            this.handleError(error);
        }
    }

    async loadExcelData() {
        try {
            this.setLoading(true);
            
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

            this.sheetNames = workbook.SheetNames;
            
            const worksheet = workbook.Sheets[workbook.SheetNames[this.currentSheet]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,
                defval: '',
                blankrows: true
            });
            
            if (jsonData.length === 0) {
                throw new Error("Het Excel-bestand is leeg of kon niet correct worden gelezen.");
            }

            this.tableData = jsonData;
            this.error = null;
        } catch (error) {
            console.error("Fout bij het ophalen of verwerken van het Excel-bestand:", error);
            this.error = `Kon de gegevens niet laden. Details: ${error.message}`;
        } finally {
            this.setLoading(false);
        }
    }

    setLoading(loading) {
        this.loading = loading;
        this.render();
    }

    handleError(error) {
        console.error("Excel Viewer Error:", error);
        this.error = error.message;
        this.loading = false;
        this.render();
    }

    async handleSheetChange(sheetIndex) {
        if (sheetIndex === this.currentSheet) return;
        
        this.currentSheet = sheetIndex;
        await this.loadExcelData();
    }

    exportToCSV() {
        try {
            const csvData = this.tableData.map(row => 
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
            
            this.showNotification('CSV export succesvol!', 'success');
        } catch (error) {
            this.showNotification('Fout bij CSV export: ' + error.message, 'error');
        }
    }

    exportToJSON() {
        try {
            const headers = this.tableData[0] || [];
            const rows = this.tableData.slice(1);
            const jsonData = rows.map(row => {
                const obj = {};
                headers.forEach((header, index) => {
                    obj[header || `Column_${index + 1}`] = row[index] || '';
                });
                return obj;
            });
            
            const blob = new Blob([JSON.stringify(jsonData, null, 2)], { type: 'application/json' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `verkeersborden_data_${new Date().toISOString().slice(0,10)}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            this.showNotification('JSON export succesvol!', 'success');
        } catch (error) {
            this.showNotification('Fout bij JSON export: ' + error.message, 'error');
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
            zIndex: '1000',
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
            setTimeout(() => document.body.removeChild(notification), 300);
        }, 3000);
    }

    setupScrollSynchronization() {
        setTimeout(() => {
            const topScroll = document.querySelector('.top-scrollbar');
            const tableContainer = document.querySelector('.table-container');
            
            if (!topScroll || !tableContainer) return;
            
            const topScrollContent = topScroll.querySelector('div');
            const dataTable = tableContainer.querySelector('.data-table');
            
            if (!topScrollContent || !dataTable) return;

            const setWidths = () => {
                topScrollContent.style.width = `${dataTable.offsetWidth}px`;
            };

            const handleTopScroll = () => {
                tableContainer.scrollLeft = topScroll.scrollLeft;
            };

            const handleTableScroll = () => {
                topScroll.scrollLeft = tableContainer.scrollLeft;
            };

            setWidths();
            topScroll.addEventListener('scroll', handleTopScroll);
            tableContainer.addEventListener('scroll', handleTableScroll);
            window.addEventListener('resize', setWidths);
            
            // Store cleanup function
            this.cleanupScrollSync = () => {
                topScroll.removeEventListener('scroll', handleTopScroll);
                tableContainer.removeEventListener('scroll', handleTableScroll);
                window.removeEventListener('resize', setWidths);
            };
        }, 100);
    }

    render() {
        const container = document.getElementById(this.containerId);
        if (!container) return;

        container.innerHTML = `
            <div class="page-wrapper">
                <div class="container">
                    <header class="header">
                        <h1 class="title">Werklijst Verkeersborden - SheetJS</h1>
                        <p class="description">
                            Modern SheetJS implementation met multi-sheet ondersteuning en geavanceerde export functionaliteit.
                        </p>
                        <div class="controls">
                            <a
                                href="${this.excelUrl}"
                                class="download-icon"
                                target="_blank"
                                rel="noopener noreferrer"
                                title="Bewerk het Excel-bestand"
                                aria-label="Download Excel bestand"
                            ></a>
                            <button onclick="excelViewer.exportToCSV()" class="export-btn" ${this.loading ? 'disabled' : ''}>
                                <i class="icon-csv"></i> Export CSV
                            </button>
                            <button onclick="excelViewer.exportToJSON()" class="export-btn" ${this.loading ? 'disabled' : ''}>
                                <i class="icon-json"></i> Export JSON
                            </button>
                        </div>
                    </header>

                    ${this.renderSheetTabs()}
                    ${this.renderContent()}
                </div>
            </div>
        `;
        
        // Re-setup scroll synchronization after render
        if (!this.loading && !this.error) {
            this.setupScrollSynchronization();
        }
    }

    renderSheetTabs() {
        if (this.sheetNames.length <= 1) return '';
        
        return `
            <div class="sheet-tabs">
                ${this.sheetNames.map((name, index) => `
                    <button
                        class="sheet-tab ${index === this.currentSheet ? 'active' : ''}"
                        onclick="excelViewer.handleSheetChange(${index})"
                        ${this.loading ? 'disabled' : ''}
                    >
                        ${this.escapeHtml(name)}
                    </button>
                `).join('')}
            </div>
        `;
    }

    renderContent() {
        if (this.loading) {
            return `
                <div class="loading-indicator">
                    <div class="spinner"></div>
                    <p>Gegevens laden...</p>
                </div>
            `;
        }
        
        if (this.error) {
            return `<div class="error-message">${this.escapeHtml(this.error)}</div>`;
        }
        
        if (this.tableData.length === 0) {
            return `<div class="empty-state">Geen gegevens beschikbaar</div>`;
        }

        return `
            <div class="stats">
                <span>📊 Rijen: ${this.tableData.length}</span>
                <span>📋 Kolommen: ${this.tableData[0] ? this.tableData[0].length : 0}</span>
                <span>📄 Werkblad: ${this.escapeHtml(this.sheetNames[this.currentSheet] || 'Onbekend')}</span>
            </div>
            <div class="top-scrollbar">
                <div></div>
            </div>
            <section class="table-container">
                <table class="data-table">
                    <thead>
                        <tr>
                            ${this.tableData[0] ? this.tableData[0].map((header, index) => `
                                <th class="table-header" data-column="${index}">
                                    ${this.escapeHtml(header || `Kolom ${index + 1}`)}
                                </th>
                            `).join('') : ''}
                        </tr>
                    </thead>
                    <tbody>
                        ${this.tableData.slice(1).map((row, rowIndex) => `
                            <tr class="table-row" data-row="${rowIndex}">
                                ${row.map((cell, cellIndex) => `
                                    <td class="table-cell" data-cell="${rowIndex}-${cellIndex}">
                                        ${this.renderCell(cell)}
                                    </td>
                                `).join('')}
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </section>
        `;
    }

    renderCell(cell) {
        if (typeof cell === "string" && cell.startsWith("http")) {
            return `
                <a
                    href="${this.escapeHtml(cell)}"
                    target="_blank"
                    rel="noopener noreferrer"
                    class="table-link"
                >
                    🔗 Bekijk link
                </a>
            `;
        }
        return this.escapeHtml(String(cell || ''));
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    destroy() {
        if (this.cleanupScrollSync) {
            this.cleanupScrollSync();
        }
    }
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', function() {
    const excelUrl = "https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1";
    window.excelViewer = new ExcelViewer('root', excelUrl);
});
