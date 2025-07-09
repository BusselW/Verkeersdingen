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
            const [loading, setLoading] = useState(false);
            const [error, setError] = useState(null);
            const [fileName, setFileName] = useState(null);
            const [searchTerm, setSearchTerm] = useState('');
            const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });
            const [currentPage, setCurrentPage] = useState(1);
            const [itemsPerPage] = useState(50);
            const fileInputRef = useRef(null);

            // SharePoint Excel file URL
            const EXCEL_URL = "https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1";

            const processExcelFile = async (file) => {
                setLoading(true);
                setError(null);
                setFileName(file.name);

                try {
                    const arrayBuffer = await file.arrayBuffer();
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(arrayBuffer);
                    const firstSheet = workbook.worksheets[0];
                    
                    const jsonData = [];
                    let headers = [];
                    
                    firstSheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
                        const rowData = [];
                        row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
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
                            
                            rowData[colNumber - 1] = cellValue;
                        });
                        
                        if (rowNumber === 1) {
                            headers = rowData;
                        } else {
                            const rowObject = {};
                            headers.forEach((header, index) => {
                                rowObject[header || `Col${index + 1}`] = rowData[index] || '';
                            });
                            jsonData.push(rowObject);
                        }
                    });
                    
                    setData(jsonData);
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
                    
                    const jsonData = [];
                    let headers = [];
                    
                    firstSheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
                        const rowData = [];
                        row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
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
                            
                            rowData[colNumber - 1] = cellValue;
                        });
                        
                        if (rowNumber === 1) {
                            headers = rowData;
                        } else {
                            const rowObject = {};
                            headers.forEach((header, index) => {
                                rowObject[header || `Col${index + 1}`] = rowData[index] || '';
                            });
                            jsonData.push(rowObject);
                        }
                    });
                    
                    setData(jsonData);
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

            const paginatedData = sortedData.slice(
                (currentPage - 1) * itemsPerPage,
                currentPage * itemsPerPage
            );

            const totalPages = Math.ceil(sortedData.length / itemsPerPage);

            const headers = data.length > 0 ? Object.keys(data[0]) : [];

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
                            <span className="data-count">{data.length} records geladen</span>
                            <button 
                                className="clear-btn"
                                onClick={() => {
                                    setData([]);
                                    setFileName(null);
                                    setError(null);
                                    setSearchTerm('');
                                    setCurrentPage(1);
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
                                        setCurrentPage(1);
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
                                    {paginatedData.map((row, rowIndex) => (
                                        <tr key={rowIndex}>
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

                    {totalPages > 1 && (
                        <div className="pagination">
                            <button 
                                onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                                disabled={currentPage === 1}
                                className="pagination-btn"
                            >
                                ← Vorige
                            </button>
                            <div className="pagination-info">
                                Pagina {currentPage} van {totalPages}
                            </div>
                            <button 
                                onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                                disabled={currentPage === totalPages}
                                className="pagination-btn"
                            >
                                Volgende →
                            </button>
                        </div>
                    )}
                </div>
            );
        };

        ReactDOM.render(<VerkeersdingendDashboard />, document.getElementById("root"));
    </script>
</body>
</html>