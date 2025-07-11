<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verkeersdingen Werklijst Dashboard</title>
    <link href="styles_new.css" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>

<body>
    <div class="dashboard">
        <header class="dashboard-header">
            <div class="header-content">
                <h1 class="dashboard-title">Verkeersdingen Werklijst</h1>
                <p class="dashboard-subtitle">Moderne werklijst viewer voor verkeersborden beheer</p>
            </div>
            
            <div class="header-actions">
                <input type="file" id="fileInput" accept=".xlsx,.xls" style="display: none;">
                <button class="upload-btn" id="reloadBtn">
                    <span class="btn-icon">&#128259;</span>
                    <span id="loadingText">Herlaad Data</span>
                </button>
                <button class="upload-btn" onclick="document.getElementById('fileInput').click()" style="margin-left: 10px; display: none;">
                    <span class="btn-icon">&#128194;</span>
                    Ander Bestand
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
    </div>

    <script>
        // Global variables like SpreadJS implementation
        let allData = [];
        let groups = [];
        let filteredData = [];
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

        // Event listeners
        document.getElementById('fileInput').addEventListener('change', handleFileUpload);
        document.getElementById('searchInput').addEventListener('input', handleSearch);
        document.getElementById('reloadBtn').addEventListener('click', loadExcelFromURL);

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
                    
                    // Use SpreadJS approach: get raw array data with header: 1
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                        header: 1,
                        defval: '',
                        blankrows: true,
                        range: 'A1:W6'  // Limit to A1:W6
                    });

                    if (jsonData.length === 0) {
                        showError('Het Excel bestand is leeg');
                        return;
                    }

                    // Extract group data from A9:C13
                    const groupData = XLSX.utils.sheet_to_json(worksheet, { 
                        header: 1,
                        defval: '',
                        blankrows: false,
                        range: 'A9:C13'  // Group definitions
                    });

                    processData(jsonData, groupData, file.name);
                    
                } catch (error) {
                    showError(`Kon het bestand niet laden: ${error.message}`);
                    showLoading(false);
                }
            };
            
            reader.readAsArrayBuffer(file);
        }

        function loadExcelFromURL() {
            const loadingText = document.getElementById('loadingText');
            const reloadBtn = document.getElementById('reloadBtn');
            
            showLoading(true);
            loadingText.textContent = 'Laden...';
            reloadBtn.disabled = true;
            
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
                    
                    // Use SpreadJS approach: get raw array data with header: 1
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                        header: 1,
                        defval: '',
                        blankrows: true,
                        range: 'A1:W6'  // Limit to A1:W6
                    });

                    // Extract group data from A9:C13
                    const groupData = XLSX.utils.sheet_to_json(worksheet, { 
                        header: 1,
                        defval: '',
                        blankrows: false,
                        range: 'A9:C13'  // Group definitions
                    });

                    processData(jsonData, groupData, 'Werklijsten MAPS PM Verkeersborden.xlsx');
                })
                .catch(error => {
                    showError(`Kon het bestand niet laden: ${error.message}`);
                    showLoading(false);
                })
                .finally(() => {
                    loadingText.textContent = 'Herlaad Data';
                    reloadBtn.disabled = false;
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
            allData = [];
            
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
                
                allData.push(rowObj);
            }

            filteredData = [...allData];
            
            // Initialize visible columns
            const allHeaders = Object.keys(allData[0] || {}).filter(key => !key.startsWith('_'));
            visibleColumns = [...allHeaders];
            
            // Update UI
            document.getElementById('fileName').innerHTML = `&#128196; <a href="${EXCEL_URL}" target="_blank" rel="noopener noreferrer" style="color: #ff6b35; text-decoration: none; font-weight: 600;">${fileName}</a>`;
            document.getElementById('dataCount').textContent = `${allData.length} records geladen (A1:W6)`;
            
            showLoading(false);
            showData();
            updateTable();
            updateGroups();
        }

        function handleSearch() {
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            
            if (allData.length === 0) return;
            
            const allHeaders = Object.keys(allData[0]).filter(key => !key.startsWith('_'));
            
            if (searchTerm === '') {
                visibleColumns = [...allHeaders];
            } else {
                visibleColumns = allHeaders.filter(header => 
                    header.toLowerCase().includes(searchTerm)
                );
            }
            
            updateTable();
            document.getElementById('resultsInfo').textContent = 
                `${visibleColumns.length} van ${allHeaders.length} kolommen`;
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

        function updateTable() {
            const thead = document.getElementById('tableHeader');
            const tbody = document.getElementById('tableBody');
            
            thead.innerHTML = '';
            tbody.innerHTML = '';
            
            if (filteredData.length === 0) return;
            
            const headerRow = document.createElement('tr');
            const allHeaders = Object.keys(filteredData[0]).filter(key => !key.startsWith('_')); // Corrected: Only one declaration
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
            
            // Show ALL data without pagination (like Excel)
            filteredData.forEach(row => {
                const tr = document.createElement('tr');
                tr.style.backgroundColor = row._groupColor || '#FFFFFF';
                
                headers.forEach((header, index) => {
                    const td = document.createElement('td');
                    const value = row[header];
                    
                    // Make column A (first column) bold - check if this is the original first column
                    // Use the allHeaders declared at the beginning of the function
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

        function showData() {
            document.getElementById('emptyState').style.display = 'none';
            document.getElementById('fileInfo').style.display = 'flex';
            document.getElementById('controlsBar').style.display = 'flex';
            document.getElementById('tableContainer').style.display = 'block';
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
                if (allData.length === 0) {
                    document.getElementById('emptyState').style.display = 'flex';
                }
            }, 5000);
        }

        function clearData() {
            allData = [];
            groups = [];
            filteredData = [];
            visibleColumns = [];
            
            document.getElementById('fileInfo').style.display = 'none';
            document.getElementById('controlsBar').style.display = 'none';
            document.getElementById('tableContainer').style.display = 'none';
            document.getElementById('groupsLegend').style.display = 'none';
            document.getElementById('emptyState').style.display = 'flex';
            document.getElementById('searchInput').value = '';
            document.getElementById('fileInput').value = '';
        }
    </script>
</body>
</html>