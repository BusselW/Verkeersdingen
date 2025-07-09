<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verkeersdingen Werklijst Dashboard - Luckysheet</title>
    <link href="styles.css" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>

<body>
    <div class="dashboard">
        <header class="dashboard-header">
            <div class="header-content">
                <h1 class="dashboard-title">Verkeersdingen Werklijst</h1>
                <p class="dashboard-subtitle">Luckysheet Implementation - Geavanceerde spreadsheet functionaliteit</p>
            </div>
            
            <div class="header-actions">
                <input type="file" id="fileInput" accept=".xlsx,.xls" style="display: none;">
                <button class="upload-btn" onclick="document.getElementById('fileInput').click()">
                    <span class="btn-icon">📂</span>
                    Excel Importeren
                </button>
            </div>
        </header>

        <div id="fileInfo" class="file-info-bar" style="display: none;">
            <span id="fileName" class="file-name"></span>
            <span id="dataCount" class="data-count"></span>
            <button class="clear-btn" onclick="clearData()">Wissen</button>
        </div>

        <div id="controlsBar" class="controls-bar" style="display: none;">
            <div class="search-container">
                <input type="text" id="searchInput" placeholder="Zoeken in alle kolommen..." class="search-input">
                <span class="search-icon">🔍</span>
            </div>
            <div id="resultsInfo" class="results-info"></div>
        </div>

        <div id="loadingState" class="loading-state" style="display: none;">
            <div class="loading-spinner"></div>
            <p>Bestand wordt verwerkt...</p>
        </div>

        <div id="errorState" class="error-state" style="display: none;">
            <div class="error-icon">⚠️</div>
            <p id="errorMessage"></p>
        </div>

        <div id="emptyState" class="empty-state">
            <div class="empty-icon">📊</div>
            <h3>Geen data geladen</h3>
            <p>Importeer een Excel bestand om aan de slag te gaan</p>
        </div>

        <div id="tableContainer" class="table-container" style="display: none;">
            <table id="dataTable" class="modern-table">
                <thead id="tableHeader"></thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>

        <div id="pagination" class="pagination" style="display: none;">
            <button id="prevBtn" class="pagination-btn">← Vorige</button>
            <div id="paginationInfo" class="pagination-info"></div>
            <button id="nextBtn" class="pagination-btn">Volgende →</button>
        </div>
    </div>

    <script>
        // Same exact JavaScript as SheetJS version
        let currentData = [];
        let filteredData = [];
        let currentPage = 1;
        const itemsPerPage = 50;
        let sortConfig = { key: null, direction: 'asc' };

        document.getElementById('fileInput').addEventListener('change', handleFileUpload);
        document.getElementById('searchInput').addEventListener('input', handleSearch);
        document.getElementById('prevBtn').addEventListener('click', () => changePage(-1));
        document.getElementById('nextBtn').addEventListener('click', () => changePage(1));

        function handleFileUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
                showError('Selecteer een geldig Excel bestand (.xlsx of .xls)');
                return;
            }

            showLoading(true);
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                        header: 1,
                        defval: '',
                        blankrows: false
                    });

                    if (jsonData.length === 0) {
                        showError('Het Excel bestand is leeg');
                        return;
                    }

                    const headers = jsonData[0];
                    const rows = jsonData.slice(1).map(row => {
                        const obj = {};
                        headers.forEach((header, index) => {
                            obj[header || `Col${index + 1}`] = row[index] || '';
                        });
                        return obj;
                    });

                    currentData = rows;
                    filteredData = [...currentData];
                    
                    document.getElementById('fileName').textContent = `📄 ${file.name}`;
                    document.getElementById('dataCount').textContent = `${currentData.length} records geladen`;
                    
                    showLoading(false);
                    showData();
                    updatePagination();
                    
                } catch (error) {
                    showError(`Kon het bestand niet laden: ${error.message}`);
                    showLoading(false);
                }
            };
            
            reader.readAsArrayBuffer(file);
        }

        function handleSearch() {
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            
            if (searchTerm === '') {
                filteredData = [...currentData];
            } else {
                filteredData = currentData.filter(row => 
                    Object.values(row).some(value => 
                        value.toString().toLowerCase().includes(searchTerm)
                    )
                );
            }
            
            currentPage = 1;
            updateTable();
            updatePagination();
        }

        function handleSort(key) {
            let direction = 'asc';
            if (sortConfig.key === key && sortConfig.direction === 'asc') {
                direction = 'desc';
            }
            sortConfig = { key, direction };

            filteredData.sort((a, b) => {
                if (a[key] < b[key]) return direction === 'asc' ? -1 : 1;
                if (a[key] > b[key]) return direction === 'asc' ? 1 : -1;
                return 0;
            });

            updateTable();
        }

        function changePage(direction) {
            const totalPages = Math.ceil(filteredData.length / itemsPerPage);
            currentPage = Math.max(1, Math.min(currentPage + direction, totalPages));
            updateTable();
            updatePagination();
        }

        function updateTable() {
            const start = (currentPage - 1) * itemsPerPage;
            const end = start + itemsPerPage;
            const pageData = filteredData.slice(start, end);
            
            const thead = document.getElementById('tableHeader');
            const tbody = document.getElementById('tableBody');
            
            thead.innerHTML = '';
            tbody.innerHTML = '';
            
            if (pageData.length === 0) return;
            
            const headerRow = document.createElement('tr');
            const headers = Object.keys(pageData[0]);
            
            headers.forEach(header => {
                const th = document.createElement('th');
                th.textContent = header;
                th.className = 'sortable';
                th.onclick = () => handleSort(header);
                
                const indicator = document.createElement('span');
                indicator.className = 'sort-indicator';
                indicator.textContent = sortConfig.key === header 
                    ? (sortConfig.direction === 'asc' ? '↑' : '↓') 
                    : '↕️';
                th.appendChild(indicator);
                
                headerRow.appendChild(th);
            });
            
            thead.appendChild(headerRow);
            
            pageData.forEach(row => {
                const tr = document.createElement('tr');
                
                headers.forEach(header => {
                    const td = document.createElement('td');
                    const value = row[header];
                    
                    if (value && value.toString().startsWith('http')) {
                        const link = document.createElement('a');
                        link.href = value;
                        link.textContent = 'Link';
                        link.className = 'table-link';
                        link.target = '_blank';
                        link.rel = 'noopener noreferrer';
                        td.appendChild(link);
                    } else {
                        td.textContent = value || '-';
                    }
                    
                    tr.appendChild(td);
                });
                
                tbody.appendChild(tr);
            });
            
            document.getElementById('resultsInfo').textContent = 
                `${filteredData.length} van ${currentData.length} records`;
        }

        function updatePagination() {
            const totalPages = Math.ceil(filteredData.length / itemsPerPage);
            
            document.getElementById('prevBtn').disabled = currentPage === 1;
            document.getElementById('nextBtn').disabled = currentPage === totalPages;
            document.getElementById('paginationInfo').textContent = 
                `Pagina ${currentPage} van ${totalPages}`;
                
            document.getElementById('pagination').style.display = totalPages > 1 ? 'flex' : 'none';
        }

        function showData() {
            document.getElementById('emptyState').style.display = 'none';
            document.getElementById('fileInfo').style.display = 'flex';
            document.getElementById('controlsBar').style.display = 'flex';
            document.getElementById('tableContainer').style.display = 'block';
            updateTable();
        }

        function showLoading(show) {
            document.getElementById('loadingState').style.display = show ? 'flex' : 'none';
        }

        function showError(message) {
            document.getElementById('errorMessage').textContent = message;
            document.getElementById('errorState').style.display = 'flex';
            document.getElementById('emptyState').style.display = 'none';
            
            setTimeout(() => {
                document.getElementById('errorState').style.display = 'none';
                if (currentData.length === 0) {
                    document.getElementById('emptyState').style.display = 'flex';
                }
            }, 5000);
        }

        function clearData() {
            currentData = [];
            filteredData = [];
            currentPage = 1;
            
            document.getElementById('fileInfo').style.display = 'none';
            document.getElementById('controlsBar').style.display = 'none';
            document.getElementById('tableContainer').style.display = 'none';
            document.getElementById('pagination').style.display = 'none';
            document.getElementById('emptyState').style.display = 'flex';
            document.getElementById('searchInput').value = '';
            document.getElementById('fileInput').value = '';
        }
    </script>
</body>
</html>