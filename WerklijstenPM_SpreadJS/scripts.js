/**
 * SpreadJS Excel Viewer - Modern Implementation
 * Enhanced with proper fallback and modern ES6+ features
 * Note: This is a fallback implementation using SheetJS since SpreadJS requires a commercial license
 */

class SpreadJSViewer {
    constructor(containerId, excelUrl) {
        this.containerId = containerId;
        this.excelUrl = excelUrl;
        this.tableData = [];
        this.sheetNames = [];
        this.currentSheet = 0;
        this.loading = true;
        this.error = null;
        this.sortColumn = null;
        this.sortDirection = 'asc';
        this.filterText = '';
        this.theme = 'light';
        
        // Bind methods
        this.handleSheetChange = this.handleSheetChange.bind(this);
        this.exportToCSV = this.exportToCSV.bind(this);
        this.exportToExcel = this.exportToExcel.bind(this);
        this.applyFormatting = this.applyFormatting.bind(this);
        this.performCalculation = this.performCalculation.bind(this);
        
        this.init();
    }

    async init() {
        try {
            await this.loadExcelData();
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
        console.error("SpreadJS Viewer Error:", error);
        this.error = error.message;
        this.loading = false;
        this.render();
    }

    async handleSheetChange(sheetIndex) {
        if (sheetIndex === this.currentSheet) return;
        
        this.currentSheet = sheetIndex;
        await this.loadExcelData();
    }

    sortTable(columnIndex) {
        if (this.tableData.length <= 1) return;
        
        const headers = this.tableData[0];
        const rows = this.tableData.slice(1);
        
        // Toggle sort direction if same column
        if (this.sortColumn === columnIndex) {
            this.sortDirection = this.sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
            this.sortColumn = columnIndex;
            this.sortDirection = 'asc';
        }
        
        rows.sort((a, b) => {
            const aVal = a[columnIndex] || '';
            const bVal = b[columnIndex] || '';
            
            // Try to convert to numbers for proper sorting
            const aNum = parseFloat(aVal);
            const bNum = parseFloat(bVal);
            
            let comparison = 0;
            if (!isNaN(aNum) && !isNaN(bNum)) {
                comparison = aNum - bNum;
            } else {
                comparison = String(aVal).localeCompare(String(bVal));
            }
            
            return this.sortDirection === 'asc' ? comparison : -comparison;
        });
        
        this.tableData = [headers, ...rows];
        this.render();
        this.showNotification(`Tabel gesorteerd op kolom ${columnIndex + 1} (${this.sortDirection})`, 'success');
    }

    filterTable(searchText) {
        this.filterText = searchText.toLowerCase();
        this.render();
    }

    getFilteredData() {
        if (!this.filterText) return this.tableData;
        
        const headers = this.tableData[0];
        const rows = this.tableData.slice(1);
        
        const filteredRows = rows.filter(row => 
            row.some(cell => 
                String(cell || '').toLowerCase().includes(this.filterText)
            )
        );
        
        return [headers, ...filteredRows];
    }

    exportToCSV() {
        try {
            const data = this.getFilteredData();
            const csvData = data.map(row => 
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

    exportToExcel() {
        try {
            const data = this.getFilteredData();
            const ws = XLSX.utils.aoa_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, this.sheetNames[this.currentSheet] || 'Sheet1');
            
            XLSX.writeFile(wb, `verkeersborden_export_${new Date().toISOString().slice(0,10)}.xlsx`);
            
            this.showNotification('Excel export succesvol!', 'success');
        } catch (error) {
            this.showNotification('Fout bij Excel export: ' + error.message, 'error');
        }
    }

    applyFormatting() {
        const tables = document.querySelectorAll('.data-table');
        tables.forEach(table => {
            // Toggle theme class
            if (this.theme === 'light') {
                table.classList.add('theme-dark');
                this.theme = 'dark';
            } else {
                table.classList.remove('theme-dark');
                this.theme = 'light';
            }
        });
        
        this.showNotification(`Thema gewijzigd naar ${this.theme}`, 'success');
    }

    performCalculation() {
        try {
            const data = this.getFilteredData();
            if (data.length <= 1) {
                this.showNotification('Geen data beschikbaar voor berekeningen', 'error');
                return;
            }
            
            const rows = data.slice(1);
            const numericColumns = [];
            
            // Find numeric columns
            for (let colIndex = 0; colIndex < (data[0] || []).length; colIndex++) {
                const values = rows.map(row => parseFloat(row[colIndex])).filter(val => !isNaN(val));
                if (values.length > 0) {
                    numericColumns.push({
                        index: colIndex,
                        name: data[0][colIndex] || `Kolom ${colIndex + 1}`,
                        values: values,
                        sum: values.reduce((a, b) => a + b, 0),
                        avg: values.reduce((a, b) => a + b, 0) / values.length,
                        min: Math.min(...values),
                        max: Math.max(...values)
                    });
                }
            }
            
            if (numericColumns.length === 0) {
                this.showNotification('Geen numerieke kolommen gevonden voor berekeningen', 'error');
                return;
            }
            
            // Show calculation results
            let message = 'Berekeningen:\n';
            numericColumns.forEach(col => {
                message += `${col.name}: Som=${col.sum.toFixed(2)}, Gem=${col.avg.toFixed(2)}, Min=${col.min}, Max=${col.max}\n`;
            });
            
            alert(message);
            this.showNotification('Berekeningen uitgevoerd', 'success');
        } catch (error) {
            this.showNotification('Fout bij berekeningen: ' + error.message, 'error');
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
            backgroundColor: type === 'success' ? '#28a745' : type === 'error' ? '#dc3545' : '#fd7e14'
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

        container.innerHTML = `
            <div class="page-wrapper">
                <div class="container">
                    <header class="header">
                        <h1 class="title">Werklijst Verkeersborden - SpreadJS</h1>
                        <p class="description">
                            Professional Excel viewer met geavanceerde functies en berekeningen (SheetJS fallback).
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
                            <button onclick="spreadJSViewer.exportToCSV()" class="export-btn" ${this.loading ? 'disabled' : ''}>
                                📄 Export CSV
                            </button>
                            <button onclick="spreadJSViewer.exportToExcel()" class="export-btn" ${this.loading ? 'disabled' : ''}>
                                📊 Export Excel
                            </button>
                            <button onclick="spreadJSViewer.applyFormatting()" class="format-btn" ${this.loading ? 'disabled' : ''}>
                                🎨 Thema
                            </button>
                            <button onclick="spreadJSViewer.performCalculation()" class="calc-btn" ${this.loading ? 'disabled' : ''}>
                                🧮 Berekenen
                            </button>
                        </div>
                    </header>

                    ${this.renderToolbar()}
                    ${this.renderSheetTabs()}
                    ${this.renderContent()}
                </div>
            </div>
        `;
    }

    renderToolbar() {
        if (this.loading || this.error) return '';
        
        return `
            <div class="toolbar">
                <div class="toolbar-section">
                    <label>🔍 Filter:</label>
                    <input 
                        type="text" 
                        placeholder="Zoek in tabel..." 
                        value="${this.filterText}"
                        onkeyup="spreadJSViewer.filterTable(this.value)"
                        style="padding: 6px 12px; border: 1px solid #ccc; border-radius: 4px; width: 200px;"
                    >
                </div>
                <div class="toolbar-section">
                    <span>📊 Thema: ${this.theme}</span>
                </div>
                <div class="toolbar-section">
                    <span>🔢 Sort: ${this.sortColumn !== null ? `Kolom ${this.sortColumn + 1} (${this.sortDirection})` : 'Geen'}</span>
                </div>
            </div>
        `;
    }

    renderSheetTabs() {
        if (this.sheetNames.length <= 1 || this.loading || this.error) return '';
        
        return `
            <div class="sheet-tabs">
                ${this.sheetNames.map((name, index) => `
                    <button
                        class="sheet-tab ${index === this.currentSheet ? 'active' : ''}"
                        onclick="spreadJSViewer.handleSheetChange(${index})"
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

        const filteredData = this.getFilteredData();
        
        return `
            <div class="stats">
                <span>📊 Totaal rijen: ${this.tableData.length}</span>
                <span>📋 Getoonde rijen: ${filteredData.length}</span>
                <span>📄 Kolommen: ${this.tableData[0] ? this.tableData[0].length : 0}</span>
                <span>📂 Werkblad: ${this.escapeHtml(this.sheetNames[this.currentSheet] || 'Onbekend')}</span>
            </div>
            <section class="table-container">
                <table class="data-table ${this.theme === 'dark' ? 'theme-dark' : ''}">
                    <thead>
                        <tr>
                            ${filteredData[0] ? filteredData[0].map((header, index) => `
                                <th class="table-header" 
                                    onclick="spreadJSViewer.sortTable(${index})"
                                    style="cursor: pointer; user-select: none;"
                                    title="Klik om te sorteren"
                                >
                                    ${this.escapeHtml(header || `Kolom ${index + 1}`)}
                                    ${this.sortColumn === index ? (this.sortDirection === 'asc' ? ' ↑' : ' ↓') : ''}
                                </th>
                            `).join('') : ''}
                        </tr>
                    </thead>
                    <tbody>
                        ${filteredData.slice(1).map((row, rowIndex) => `
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
        
        // Check if cell is numeric for special formatting
        const numValue = parseFloat(cell);
        if (!isNaN(numValue) && cell !== '') {
            return `<span class="numeric-cell">${this.escapeHtml(String(cell))}</span>`;
        }
        
        return this.escapeHtml(String(cell || ''));
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
}

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', function() {
    const excelUrl = "https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1";
    window.spreadJSViewer = new SpreadJSViewer('root', excelUrl);
});
