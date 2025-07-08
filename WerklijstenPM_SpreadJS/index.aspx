<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="Geavanceerde Excel viewer met SheetJS fallback en extra functies zoals sorteren, filteren en berekeningen">
    <title>Verkeersborden Werklijst - SpreadJS Fallback</title>
    <link href="styles.css" rel="stylesheet">
    <link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>📊</text></svg>">
    
    <!-- SheetJS as fallback since SpreadJS requires commercial license -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>

<body>
    <div id="root"></div>
    <script src="scripts.js"></script>
</body>
</html>


        const VerkeersbordenWerklijst = () => {
            const [tableData, setTableData] = useState([]);
            const [loading, setLoading] = useState(true);
            const [error, setError] = useState(null);
            const [sheetNames, setSheetNames] = useState([]);
            const [currentSheet, setCurrentSheet] = useState(0);
            const [zoom, setZoom] = useState(1);
            const [theme, setTheme] = useState('default');

            // Refs voor de scrollbars
            const topScrollRef = useRef(null);
            const tableContainerRef = useRef(null);

            useEffect(() => {
                const fetchExcelData = async () => {
                    try {
                        const response = await fetch("https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1");
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

                        // Store all sheet names
                        setSheetNames(workbook.SheetNames);
                        
                        // Process the current sheet
                        const worksheet = workbook.Sheets[workbook.SheetNames[currentSheet]];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                            header: 1,
                            defval: '',
                            blankrows: true
                        });
                        
                        if(jsonData.length === 0){
                            throw new Error("Het Excel-bestand is leeg of kon niet correct worden gelezen.");
                        }

                        setTableData(jsonData);
                    } catch (e) {
                        console.error("Fout bij het ophalen of verwerken van het Excel-bestand:", e);
                        setError(`Kon de gegevens niet laden. Details: ${e.message}`);
                    } finally {
                        setLoading(false);
                    }
                };

                fetchExcelData();
            }, [currentSheet]);
            
            // Effect voor het synchroniseren van de scrollbars
            useEffect(() => {
                const topScroll = topScrollRef.current;
                const tableContainer = tableContainerRef.current;
                
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

                return () => {
                    topScroll.removeEventListener('scroll', handleTopScroll);
                    tableContainer.removeEventListener('scroll', handleTableScroll);
                    window.removeEventListener('resize', setWidths);
                };
            }, [loading]);

            const handleSheetChange = (sheetIndex) => {
                setCurrentSheet(sheetIndex);
                setLoading(true);
            };

            const exportToExcel = () => {
                try {
                    const ws = XLSX.utils.aoa_to_sheet(tableData);
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Verkeersborden");
                    XLSX.writeFile(wb, "Verkeersborden_Export.xlsx");
                    showSuccess('Excel bestand wordt gedownload...');
                } catch (error) {
                    showError('Fout bij exporteren: ' + error.message);
                }
            };

            const autoFormat = () => {
                showSuccess('Auto-formattering toegepast! (Demo functie)');
            };

            const recalculate = () => {
                showSuccess('Herberekening voltooid! (Demo functie)');
            };

            const changeZoom = (newZoom) => {
                setZoom(newZoom);
            };

            const changeTheme = (newTheme) => {
                setTheme(newTheme);
            };

            const showError = (message) => {
                console.error(message);
                alert(message);
            };

            const showSuccess = (message) => {
                console.log(message);
                // Create temporary success message
                const successDiv = document.createElement('div');
                successDiv.textContent = message;
                successDiv.style.cssText = `
                    position: fixed;
                    top: 20px;
                    right: 20px;
                    background: #28a745;
                    color: white;
                    padding: 15px 20px;
                    border-radius: 5px;
                    z-index: 10000;
                    box-shadow: 0 4px 8px rgba(0,0,0,0.2);
                `;
                document.body.appendChild(successDiv);
                
                setTimeout(() => {
                    if (document.body.contains(successDiv)) {
                        document.body.removeChild(successDiv);
                    }
                }, 3000);
            };

            return (
                <div className="page-wrapper">
                    <div className="container">
                        <header className="header">
                            <h1 className="title">Werklijst Verkeersborden - Enterprise Edition</h1>
                            <p className="description">
                                Enterprise-grade Excel viewer met geavanceerde functies. (Gebruikt SheetJS als fallback)
                            </p>
                            <div className="controls">
                                <a
                                    href="https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1"
                                    className="download-icon"
                                    target="_blank"
                                    rel="noopener noreferrer"
                                    title="Bewerk het Excel-bestand"
                                ></a>
                                <button onClick={() => window.location.reload()} className="load-btn">Herlaad Data</button>
                                <button onClick={exportToExcel} className="export-btn">Export Excel</button>
                                <button onClick={autoFormat} className="format-btn">Auto Format</button>
                                <button onClick={recalculate} className="calc-btn">Herbereken</button>
                            </div>
                        </header>
                        
                        <div className="toolbar">
                            <div className="toolbar-section">
                                <label>Zoom:</label>
                                <select value={zoom} onChange={(e) => changeZoom(parseFloat(e.target.value))}>
                                    <option value="0.5">50%</option>
                                    <option value="0.75">75%</option>
                                    <option value="1">100%</option>
                                    <option value="1.25">125%</option>
                                    <option value="1.5">150%</option>
                                    <option value="2">200%</option>
                                </select>
                            </div>
                            <div className="toolbar-section">
                                <label>Thema:</label>
                                <select value={theme} onChange={(e) => changeTheme(e.target.value)}>
                                    <option value="default">Default</option>
                                    <option value="dark">Dark</option>
                                    <option value="blue">Blue</option>
                                    <option value="green">Green</option>
                                </select>
                            </div>
                            <div className="toolbar-section">
                                <span>Cel: A1</span>
                                <span>Sheet: {sheetNames[currentSheet] || 'Sheet1'}</span>
                            </div>
                        </div>

                        {sheetNames.length > 1 && (
                            <div className="sheet-tabs">
                                {sheetNames.map((name, index) => (
                                    <button
                                        key={index}
                                        className={`sheet-tab ${index === currentSheet ? 'active' : ''}`}
                                        onClick={() => handleSheetChange(index)}
                                    >
                                        {name}
                                    </button>
                                ))}
                            </div>
                        )}
                        
                        {loading && <div className="loading-indicator">Spreadsheet wordt geladen...</div>}
                        {error && <div className="error-message">{error}</div>}

                        {!loading && !error && (
                            <React.Fragment>
                                <div className="stats">
                                    Rijen: {tableData.length} | Kolommen: {tableData[0] ? tableData[0].length : 0}
                                </div>
                                <div className="top-scrollbar" ref={topScrollRef}>
                                    <div></div>
                                </div>
                                <section 
                                    className="table-container" 
                                    ref={tableContainerRef}
                                    style={{ 
                                        transform: `scale(${zoom})`,
                                        transformOrigin: 'top left',
                                        width: `${100/zoom}%`,
                                        height: `${100/zoom}%`
                                    }}
                                >
                                    <table className={`data-table theme-${theme}`}>
                                        <thead>
                                            <tr>
                                                {tableData[0] &&
                                                    tableData[0].map((header, index) => (
                                                        <th key={index} className="table-header">
                                                            {header || `Kolom ${index + 1}`}
                                                        </th>
                                                    ))}
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {tableData.slice(1).map((row, rowIndex) => (
                                                <tr key={rowIndex} className="table-row">
                                                    {row.map((cell, cellIndex) => (
                                                        <td key={cellIndex} className="table-cell">
                                                            {typeof cell === "string" && cell.startsWith("http") ? (
                                                                <a
                                                                    href={cell}
                                                                    target="_blank"
                                                                    rel="noopener noreferrer"
                                                                    className="table-link"
                                                                >
                                                                    Bekijk link
                                                                </a>
                                                            ) : (
                                                                cell
                                                            )}
                                                        </td>
                                                    ))}
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </section>
                            </React.Fragment>
                        )}
                    </div>
                </div>
            );
        };

        ReactDOM.render(<VerkeersbordenWerklijst />, document.getElementById("root"));
    </script>
</body>
</html>
