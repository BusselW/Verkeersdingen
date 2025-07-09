<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verkeersdingen Werklijst Dashboard - SpreadJS</title>
    <link href="styles.css" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>

<body>
    <div class="dashboard">
        <header class="dashboard-header">
            <div class="header-content">
                <h1 class="dashboard-title">Verkeersdingen Werklijst</h1>
                <p class="dashboard-subtitle">SpreadJS Implementation - Enterprise spreadsheet oplossing</p>
            </div>
            
            <div class="header-actions">
                <input type="file" id="fileInput" accept=".xlsx,.xls" style="display: none;">
                <button class="upload-btn" onclick="document.getElementById('fileInput').click()" style="display: none;">
                    <span class="btn-icon">&#128194;</span>
                    Excel Importeren
                </button>
            </div>
        </header>

        <div id="fileInfo" class="file-info-bar" style="display: none;">
            <span id="fileName" class="file-name"></span>
            <span id="dataCount" class="data-count"></span>
            <button class="clear-btn" onclick="clearData()" style="display: none;">Wissen</button>
        </div>

        <div id="controlsBar" class="controls-bar" style="display: none;">
            <div class="search-container">
                <input type="text" id="searchInput" placeholder="Filter kolommen..." class="search-input">
                <span class="search-icon">&#128269;</span>
            </div>
            <div id="resultsInfo" class="results-info"></div>
        </div>

        <div id="loadingState" class="loading-state" style="display: none;">
            <div class="loading-spinner"></div>
            <p>Bestand wordt verwerkt...</p>
        </div>

        <div id="errorState" class="error-state" style="display: none;">
            <div class="error-icon">&#9888;</div>
            <p id="errorMessage"></p>
        </div>

        <div id="emptyState" class="empty-state">
            <div class="empty-icon">&#128202;</div>
            <h3>Geen data geladen</h3>
            <p>Importeer een Excel bestand om aan de slag te gaan</p>
        </div>

        <div id="tableContainer" class="table-container" style="display: none;">
            <table id="dataTable" class="modern-table">
                <thead id="tableHeader"></thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>

        <div id="groupsLegend" class="groups-legend" style="display: none;">
            <h3 class="legend-title">Groepen Overzicht</h3>
            <div id="groupsGrid" class="groups-grid">
                <!-- Groups will be populated here -->
            </div>
        </div>

        <div id="pagination" class="pagination" style="display: none;">
            <button id="prevBtn" class="pagination-btn">← Vorige</button>
            <div id="paginationInfo" class="pagination-info"></div>
            <button id="nextBtn" class="pagination-btn">Volgende →</button>
        </div>
    </div>

    <script>
        // Same JavaScript implementation as other versions
        let currentData = [];
        let filteredData = [];
        let groups = [];
        let currentPage = 1;
        const itemsPerPage = 50;
        let sortConfig = { key: null, direction: 'asc' };
        let visibleColumns = []; // Track which columns should be visible

        // Professional muted colors for groups
        const groupColors = {
            'A': '#f8fafc', // Light gray
            'B': '#f1f5f9', // Slate
            'C': '#f0f9ff', // Light blue
            'D': '#f0fdf4', // Light green
            'E': '#fefce8', // Light yellow
        };

        // SharePoint Excel file URL
        const EXCEL_URL = "https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1";

        document.getElementById('fileInput').addEventListener('change', handleFileUpload);
        document.getElementById('searchInput').addEventListener('input', handleSearch);
        document.getElementById('prevBtn').addEventListener('click', () => changePage(-1));
        document.getElementById('nextBtn').addEventListener('click', () => changePage(1));

        // Auto-load Excel file from SharePoint on page load
        window.addEventListener('load', function() {
            loadExcelFromURL();
        });

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
                    
                    // Get main data from A1:W6
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                        header: 1,
                        defval: '',
                        blankrows: true,
                        range: 'A1:W6'
                    });

                    // Get group data from A9:C13
                    const groupData = XLSX.utils.sheet_to_json(worksheet, { 
                        header: 1,
                        defval: '',
                        blankrows: false,
                        range: 'A9:C13'
                    });

                    if (jsonData.length === 0) {
                        showError('Het Excel bestand is leeg');
                        return;
                    }

                    processData(jsonData, groupData, file.name);
                    
                } catch (error) {
                    showError(`Kon het bestand niet laden: ${error.message}`);
                    showLoading(false);
                }
            };
            
            reader.readAsArrayBuffer(file);
        }

        function handleSearch() {
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            
            if (currentData.length === 0) return;
            
            const allHeaders = Object.keys(currentData[0]).filter(key => !key.startsWith('_'));
            
            if (searchTerm === '') {
                visibleColumns = [...allHeaders];
            } else {
                visibleColumns = allHeaders.filter(header => 
                    header.toLowerCase().includes(searchTerm)
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
            const allHeaders = Object.keys(pageData[0]).filter(key => !key.startsWith('_'));
            const headers = visibleColumns.length > 0 ? visibleColumns : allHeaders;
            
            headers.forEach((header, index) => {
                const th = document.createElement('th');
                th.textContent = header;
                th.className = 'sortable';
                th.onclick = () => handleSort(header);
                
                // Merge cells D1:W1 (columns 3 to end, assuming 0-indexed) - only if we have enough columns
                const originalIndex = allHeaders.indexOf(header);
                if (originalIndex === 3 && headers.length - index > 1) {
                    th.colSpan = headers.length - index; // Span from column D to the end
                    th.textContent = 'WERKLIJSTEN';
                    th.style.textAlign = 'center';
                    th.style.fontWeight = 'bold';
                    th.onclick = null; // Remove sorting for merged header
                } else if (originalIndex > 3 && allHeaders.indexOf(headers[index-1]) === 3) {
                    return; // Skip remaining headers as they're merged
                }
                
                const indicator = document.createElement('span');
                indicator.className = 'sort-indicator';
                if (originalIndex !== 3 || headers.length - index === 1) { // Don't add sort indicator to merged cell
                    indicator.innerHTML = sortConfig.key === header 
                        ? (sortConfig.direction === 'asc' ? '&#8593;' : '&#8595;') 
                        : '&#8597;';
                    th.appendChild(indicator);
                }
                
                headerRow.appendChild(th);
            });
            
            thead.appendChild(headerRow);
            
            pageData.forEach(row => {
                const tr = document.createElement('tr');
                tr.style.backgroundColor = row._groupColor || '#FFFFFF';
                
                headers.forEach((header, index) => {
                    const td = document.createElement('td');
                    const value = row[header];
                    
                    // Make column A (first column) bold - check if this is the original first column
                    const allHeaders = Object.keys(row).filter(key => !key.startsWith('_'));
                    const originalIndex = allHeaders.indexOf(header);
                    if (originalIndex === 0) {
                        td.style.fontWeight = 'bold';
                    }
                    
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
            
            const allHeaders = Object.keys(currentData[0] || {}).filter(key => !key.startsWith('_'));
            document.getElementById('resultsInfo').textContent = 
                `${headers.length} van ${allHeaders.length} kolommen`;
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

        function loadExcelFromURL() {
            showLoading(true);
            
            fetch(EXCEL_URL)
                .then(response => {
                    if (!response.ok) {
                        throw new Error(`Het ophalen van het bestand is mislukt met status: ${response.status}`);
                    }
                    return response.arrayBuffer();
                })
                .then(arrayBuffer => {
                    const data = new Uint8Array(arrayBuffer);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // Get main data from A1:W6
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                        header: 1,
                        defval: '',
                        blankrows: true,
                        range: 'A1:W6'
                    });

                    // Get group data from A9:C13
                    const groupData = XLSX.utils.sheet_to_json(worksheet, { 
                        header: 1,
                        defval: '',
                        blankrows: false,
                        range: 'A9:C13'
                    });

                    if (jsonData.length === 0) {
                        showError('Het Excel bestand is leeg');
                        return;
                    }

                    processData(jsonData, groupData, 'Werklijsten MAPS PM Verkeersborden.xlsx');
                })
                .catch(error => {
                    showError(`Kon het bestand niet laden: ${error.message}`);
                    showLoading(false);
                });
        }

        function processData(jsonData, groupData, fileName) {
            // Process groups first (A9:C13)
            groups = [];
            groupData.forEach(row => {
                if (row[0] && row[0].toString().trim()) {
                    const groupLetter = row[0].toString().trim().slice(-1); // Get last character
                    groups.push({
                        letter: groupLetter,
                        name: row[0] || '',
                        members: row[1] || '',
                        contact: row[2] || '',
                        color: groupColors[groupLetter] || '#FFFFFF'
                    });
                }
            });

            // Process main data (A1:W6)
            const headers = jsonData[0] || [];
            currentData = [];
            
            // Process data rows (rows 2-6, which are indices 1-5 in the array)
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                const rowObj = {};
                
                headers.forEach((header, index) => {
                    rowObj[header || `Col${index + 1}`] = row[index] || '';
                });
                
                // Match group colors based on column A value (first column)
                const columnAValue = row[0] ? row[0].toString().trim() : '';
                const matchingGroup = groups.find(group => group.letter === columnAValue);
                rowObj._groupColor = matchingGroup ? matchingGroup.color : '#FFFFFF';
                rowObj._groupLetter = columnAValue;
                
                currentData.push(rowObj);
            }

            filteredData = [...currentData];
            
            // Initialize visible columns
            const allHeaders = Object.keys(currentData[0] || {}).filter(key => !key.startsWith('_'));
            visibleColumns = [...allHeaders];
            
            // Update UI
            document.getElementById('fileName').innerHTML = `&#128196; <a href="${EXCEL_URL}" target="_blank" rel="noopener noreferrer" style="color: #ff6b35; text-decoration: none; font-weight: 600;">${fileName}</a>`;
            document.getElementById('dataCount').textContent = `${currentData.length} records geladen (A1:W6)`;
            
            showLoading(false);
            showData();
            updatePagination();
            updateGroups();
        }

        function updateGroups() {
            const groupsGrid = document.getElementById('groupsGrid');
            groupsGrid.innerHTML = '';
            
            groups.forEach(group => {
                const groupCard = document.createElement('div');
                groupCard.className = 'group-card';
                groupCard.style.backgroundColor = group.color;
                
                groupCard.innerHTML = `
                    <div class="group-header">
                        <span class="group-name">${group.name}</span>
                        <span class="group-letter">${group.letter}</span>
                    </div>
                    <div class="group-details">
                        <div class="group-row">
                            <strong>Leden:</strong> ${group.members}
                        </div>
                        <div class="group-row">
                            ${group.contact}
                        </div>
                    </div>
                `;
                
                groupsGrid.appendChild(groupCard);
            });
            
            document.getElementById('groupsLegend').style.display = groups.length > 0 ? 'block' : 'none';
        }

        function clearData() {
            currentData = [];
            filteredData = [];
            groups = [];
            currentPage = 1;
            visibleColumns = [];
            
            document.getElementById('fileInfo').style.display = 'none';
            document.getElementById('controlsBar').style.display = 'none';
            document.getElementById('tableContainer').style.display = 'none';
            document.getElementById('groupsLegend').style.display = 'none';
            document.getElementById('pagination').style.display = 'none';
            document.getElementById('emptyState').style.display = 'flex';
            document.getElementById('searchInput').value = '';
            document.getElementById('fileInput').value = '';
        }
    </script>
</body>
</html>