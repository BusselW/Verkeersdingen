<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verkeersdingen Werklijst Dashboard</title>
    <link href="styles_new.css" rel="stylesheet">
    <script src="https://unpkg.com/react@17/umd/react.production.min.js" crossorigin></script>
    <script src="https://unpkg.com/react-dom@17/umd/react-dom.production.min.js" crossorigin></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/6.26.0/babel.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/exceljs@4.3.0/dist/exceljs.min.js"></script>
</head>

<body>
    <div id="root"></div>
    <script type="text/babel">
        const { useState, useEffect, useRef } = React;

        const VerkeersdingendDashboard = () => {
            const [data, setData] = useState([]);
            const [groups, setGroups] = useState([]);
            const [loading, setLoading] = useState(false);
            const [error, setError] = useState(null);
            const [fileName, setFileName] = useState(null);
            const [searchTerm, setSearchTerm] = useState('');
            const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });
            const fileInputRef = useRef(null);

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

            const processExcelDataWithGroups = (worksheet) => {
                // Extract main data from A1:W6
                const mainData = [];
                const headers = [];
                
                // Get headers from row 1 (A1:W1)
                for (let col = 1; col <= 23; col++) { // A=1 to W=23
                    const cell = worksheet.getCell(1, col);
                    let cellValue = cell.value || '';
                    
                    if (cell.value && typeof cell.value === 'object') {
                        if (cell.value.richText) {
                            cellValue = cell.value.richText.map(rt => rt.text).join('');
                        } else if (cell.value.formula) {
                            cellValue = cell.value.result || cell.value.formula;
                        } else if (cell.value.hyperlink) {
                            cellValue = cell.value.text || cell.value.hyperlink;
                        } else if (cell.value instanceof Date) {
                            cellValue = cell.value.toLocaleDateString('nl-NL');
                        } else {
                            cellValue = cell.value.toString();
                        }
                    } else if (cell.value instanceof Date) {
                        cellValue = cell.value.toLocaleDateString('nl-NL');
                    }
                    
                    headers.push(cellValue || `Col${col}`);
                }
                
                // Get data rows from A2:W6
                for (let row = 2; row <= 6; row++) {
                    const rowData = {};
                    for (let col = 1; col <= 23; col++) {
                        const cell = worksheet.getCell(row, col);
                        let cellValue = cell.value || '';
                        
                        if (cell.value && typeof cell.value === 'object') {
                            if (cell.value.richText) {
                                cellValue = cell.value.richText.map(rt => rt.text).join('');
                            } else if (cell.value.formula) {
                                cellValue = cell.value.result || cell.value.formula;
                            } else if (cell.value.hyperlink) {
                                cellValue = cell.value.text || cell.value.hyperlink;
                            } else if (cell.value instanceof Date) {
                                cellValue = cell.value.toLocaleDateString('nl-NL');
                            } else {
                                cellValue = cell.value.toString();
                            }
                        } else if (cell.value instanceof Date) {
                            cellValue = cell.value.toLocaleDateString('nl-NL');
                        }
                        
                        rowData[headers[col - 1]] = cellValue;
                    }
                    
                    mainData.push(rowData);
                }
                
                // Extract group definitions from A9:C13
                const groupDefinitions = [];
                for (let row = 9; row <= 13; row++) {
                    const groepjeCell = worksheet.getCell(row, 1); // Column A
                    const membersCell = worksheet.getCell(row, 2); // Column B
                    const contactCell = worksheet.getCell(row, 3); // Column C
                    
                    let groepje = groepjeCell.value || '';
                    let members = membersCell.value || '';
                    let contact = contactCell.value || '';
                    
                    // Process cell values
                    [groepje, members, contact].forEach((cell, index) => {
                        if (cell && typeof cell === 'object') {
                            if (cell.richText) {
                                cell = cell.richText.map(rt => rt.text).join('');
                            } else if (cell.formula) {
                                cell = cell.result || cell.formula;
                            } else if (cell.hyperlink) {
                                cell = cell.text || cell.hyperlink;
                            } else if (cell instanceof Date) {
                                cell = cell.toLocaleDateString('nl-NL');
                            } else {
                                cell = cell.toString();
                            }
                        }
                        
                        if (index === 0) groepje = cell;
                        if (index === 1) members = cell;
                        if (index === 2) contact = cell;
                    });
                    
                    // Trim to get last character (e.g., "Groepje A" -> "A")
                    const groupLetter = groepje.toString().trim().slice(-1);
                    
                    if (groupLetter && groepje.toString().trim()) {
                        groupDefinitions.push({
                            letter: groupLetter,
                            name: groepje.toString(),
                            members: members.toString(),
                            contact: contact.toString(),
                            color: groupColors[groupLetter] || '#FFFFFF'
                        });
                    }
                }
                
                // Now match group letters against A4:A6 values and assign colors
                mainData.forEach(row => {
                    const columnAValue = row[headers[0]] ? row[headers[0]].toString().trim() : '';
                    const matchingGroup = groupDefinitions.find(group => group.letter === columnAValue);
                    
                    row._groupLetter = columnAValue;
                    row._groupColor = matchingGroup ? matchingGroup.color : '#FFFFFF';
                });
                
                return { mainData, groupDefinitions };
            };

            const processExcelFile = async (file) => {
                setLoading(true);
                setError(null);
                setFileName(file.name);

                try {
                    const arrayBuffer = await file.arrayBuffer();
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(arrayBuffer);
                    const firstSheet = workbook.worksheets[0];
                    
                    const { mainData, groupDefinitions } = processExcelDataWithGroups(firstSheet);
                    
                    setData(mainData);
                    setGroups(groupDefinitions);
                } catch (e) {
                    setError(`Kon het bestand niet laden: ${e.message}`);
                } finally {
                    setLoading(false);
                }
            };

            const loadExcelFromURL = async () => {
                setLoading(true);
                setError(null);
                setFileName('Werklijsten MAPS PM Verkeersborden.xlsx');

                try {
                    const response = await fetch(EXCEL_URL);
                    if (!response.ok) {
                        throw new Error(`Het ophalen van het bestand is mislukt met status: ${response.status}`);
                    }
                    
                    const arrayBuffer = await response.arrayBuffer();
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(arrayBuffer);
                    const firstSheet = workbook.worksheets[0];
                    
                    const { mainData, groupDefinitions } = processExcelDataWithGroups(firstSheet);
                    
                    setData(mainData);
                    setGroups(groupDefinitions);
                } catch (e) {
                    setError(`Kon het bestand niet laden: ${e.message}`);
                } finally {
                    setLoading(false);
                }
            };

            const handleFileUpload = (event) => {
                const file = event.target.files[0];
                if (file) {
                    if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                        file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                        processExcelFile(file);
                    } else {
                        setError('Selecteer een geldig Excel bestand (.xlsx of .xls)');
                    }
                }
            };

            const handleSort = (key) => {
                let direction = 'asc';
                if (sortConfig.key === key && sortConfig.direction === 'asc') {
                    direction = 'desc';
                }
                setSortConfig({ key, direction });
            };

            const filteredData = data.filter(item =>
                Object.values(item).some(value =>
                    value.toString().toLowerCase().includes(searchTerm.toLowerCase())
                )
            );

            const sortedData = React.useMemo(() => {
                let sortableItems = [...filteredData];
                if (sortConfig.key) {
                    sortableItems.sort((a, b) => {
                        if (a[sortConfig.key] < b[sortConfig.key]) {
                            return sortConfig.direction === 'asc' ? -1 : 1;
                        }
                        if (a[sortConfig.key] > b[sortConfig.key]) {
                            return sortConfig.direction === 'asc' ? 1 : -1;
                        }
                        return 0;
                    });
                }
                return sortableItems;
            }, [filteredData, sortConfig]);


            const headers = data.length > 0 ? Object.keys(data[0]).filter(key => !key.startsWith('_')) : [];

            // Auto-load Excel file from SharePoint on component mount
            useEffect(() => {
                loadExcelFromURL();
            }, []);

            return (
                <div className="dashboard">
                    <header className="dashboard-header">
                        <div className="header-content">
                            <h1 className="dashboard-title">Verkeersdingen Werklijst</h1>
                            <p className="dashboard-subtitle">Moderne werklijst viewer voor verkeersborden beheer</p>
                        </div>
                        
                        <div className="header-actions">
                            <input
                                ref={fileInputRef}
                                type="file"
                                accept=".xlsx,.xls"
                                onChange={handleFileUpload}
                                style={{ display: 'none' }}
                            />
                            <button 
                                className="upload-btn"
                                onClick={() => loadExcelFromURL()}
                                disabled={loading}
                            >
                                <span className="btn-icon">🔄</span>
                                {loading ? 'Laden...' : 'Herlaad Data'}
                            </button>
                            <button 
                                className="upload-btn"
                                onClick={() => fileInputRef.current && fileInputRef.current.click()}
                                style={{ marginLeft: '10px' }}
                            >
                                <span className="btn-icon">📂</span>
                                Ander Bestand
                            </button>
                        </div>
                    </header>

                    {fileName && (
                        <div className="file-info-bar">
                            <span className="file-name">📄 {fileName}</span>
                            <span className="data-count">{data.length} records geladen (A1:W6)</span>
                            <button 
                                className="clear-btn"
                                onClick={() => {
                                    setData([]);
                                    setGroups([]);
                                    setFileName(null);
                                    setError(null);
                                    setSearchTerm('');
                                }}
                            >
                                Wissen
                            </button>
                        </div>
                    )}

                    {data.length > 0 && (
                        <div className="controls-bar">
                            <div className="search-container">
                                <input
                                    type="text"
                                    placeholder="Zoeken in alle kolommen..."
                                    value={searchTerm}
                                    onChange={(e) => {
                                        setSearchTerm(e.target.value);
                                    }}
                                    className="search-input"
                                />
                                <span className="search-icon">🔍</span>
                            </div>
                            <div className="results-info">
                                {filteredData.length} van {data.length} records
                            </div>
                        </div>
                    )}

                    {loading && (
                        <div className="loading-state">
                            <div className="loading-spinner"></div>
                            <p>Bestand wordt verwerkt...</p>
                        </div>
                    )}

                    {error && (
                        <div className="error-state">
                            <div className="error-icon">⚠️</div>
                            <p>{error}</p>
                        </div>
                    )}

                    {!loading && !error && data.length === 0 && (
                        <div className="empty-state">
                            <div className="empty-icon">📊</div>
                            <h3>Geen data geladen</h3>
                            <p>Importeer een Excel bestand om aan de slag te gaan</p>
                        </div>
                    )}

                    {!loading && !error && data.length > 0 && (
                        <div className="table-container">
                            <table className="modern-table">
                                <thead>
                                    <tr>
                                        {headers.map((header, index) => (
                                            <th 
                                                key={index}
                                                onClick={() => handleSort(header)}
                                                className={`sortable ${sortConfig.key === header ? sortConfig.direction : ''}`}
                                            >
                                                {header}
                                                <span className="sort-indicator">
                                                    {sortConfig.key === header 
                                                        ? (sortConfig.direction === 'asc' ? '↑' : '↓') 
                                                        : '↕️'}
                                                </span>
                                            </th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {sortedData.map((row, rowIndex) => (
                                        <tr key={rowIndex} style={{ backgroundColor: row._groupColor || '#FFFFFF' }}>
                                            {headers.map((header, colIndex) => (
                                                <td key={colIndex}>
                                                    {row[header] && row[header].toString().startsWith('http') ? (
                                                        <a 
                                                            href={row[header]} 
                                                            target="_blank" 
                                                            rel="noopener noreferrer"
                                                            className="table-link"
                                                        >
                                                            Link
                                                        </a>
                                                    ) : (
                                                        row[header] || '-'
                                                    )}
                                                </td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}

                    {!loading && !error && groups.length > 0 && (
                        <div className="groups-legend">
                            <h3 className="legend-title">Groepen Overzicht</h3>
                            <div className="groups-grid">
                                {groups.map((group, index) => (
                                    <div key={index} className="group-card" style={{ backgroundColor: group.color }}>
                                        <div className="group-header">
                                            <span className="group-name">{group.name}</span>
                                            <span className="group-letter">{group.letter}</span>
                                        </div>
                                        <div className="group-details">
                                            <div className="group-row">
                                                <strong>Leden:</strong> {group.members}
                                            </div>
                                            <div className="group-row">
                                                <strong>Contactpersoon:</strong> {group.contact}
                                            </div>
                                        </div>
                                    </div>
                                ))}
                            </div>
                        </div>
                    )}

                </div>
            );
        };

        ReactDOM.render(<VerkeersdingendDashboard />, document.getElementById("root"));
    </script>
</body>
</html>