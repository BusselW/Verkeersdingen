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
                <button class="upload-btn" onclick="document.getElementById('fileInput').click()">
                    <span class="btn-icon">&#128194;</span>
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

        // Pastel colors for groups
        const groupColors = {
            'A': '#FFE5E5', // Light red
            'B': '#E5F3FF', // Light blue
            'C': '#E5FFE5', // Light green
            'D': '#FFF5E5', // Light orange
            'E': '#F0E5FF', // Light purple
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
            const headers = Object.keys(pageData[0]).filter(key => !key.startsWith('_'));
            
            headers.forEach(header => {
                const th = document.createElement('th');
                th.textContent = header;
                th.className = 'sortable';
                th.onclick = () => handleSort(header);
                
                const indicator = document.createElement('span');
                indicator.className = 'sort-indicator';
                indicator.innerHTML = sortConfig.key === header 
                    ? (sortConfig.direction === 'asc' ? '&#8593;' : '&#8595;') 
                    : '&#8597;';
                th.appendChild(indicator);
                
                headerRow.appendChild(th);
            });
            
            thead.appendChild(headerRow);
            
            pageData.forEach(row => {
                const tr = document.createElement('tr');
                tr.style.backgroundColor = row._groupColor || '#FFFFFF';
                
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
            
            // Update UI
            document.getElementById('fileName').innerHTML = `&#128196; ${fileName}`;
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
                            <strong>Contactpersoon:</strong> ${group.contact}
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