<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verkeersdingen Werklijst - Excel Viewer</title>
    <link href="styles_new.css" rel="stylesheet">
    <script src="https://unpkg.com/react@17/umd/react.production.min.js" crossorigin></script>
    <script src="https://unpkg.com/react-dom@17/umd/react-dom.production.min.js" crossorigin></script>
    <!-- Note: Using Babel in-browser for development. For production, consider precompiling JSX -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/6.26.0/babel.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/exceljs@4.3.0/dist/exceljs.min.js"></script>
</head>

<body>
    <div id="root"></div>
    <script type="text/babel">
        const { useState, useEffect, useRef } = React;

        // Error Boundary Component
        class ErrorBoundary extends React.Component {
            constructor(props) {
                super(props);
                this.state = { hasError: false, error: null };
            }

            static getDerivedStateFromError(error) {
                return { hasError: true, error: error };
            }

            componentDidCatch(error, errorInfo) {
                console.error('Error Boundary caught an error:', error, errorInfo);
            }

            render() {
                if (this.state.hasError) {
                    return (
                        <div className="error-message">
                            <h2>Er is een fout opgetreden</h2>
                            <p>De applicatie heeft een onverwachte fout ondervonden.</p>
                            <button onClick={() => window.location.reload()}>
                                Pagina herladen
                            </button>
                        </div>
                    );
                }

                return this.props.children;
            }
        }

        const VerkeersbordenWerklijst = () => {
            const [tableData, setTableData] = useState([]);
            const [loading, setLoading] = useState(false);
            const [error, setError] = useState(null);
            const [fileName, setFileName] = useState(null);
            const fileInputRef = useRef(null);

            // Refs voor de scrollbars
            const topScrollRef = useRef(null);
            const tableContainerRef = useRef(null);

            const processExcelFile = async (file) => {
                setLoading(true);
                setError(null);
                setFileName(file.name);

                try {
                    const arrayBuffer = await file.arrayBuffer();
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(arrayBuffer);
                    const firstSheet = workbook.worksheets[0];
                    
                    // Convert ExcelJS worksheet to array format
                    const jsonData = [];
                    try {
                        firstSheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
                            const rowData = [];
                            row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
                                let cellValue = cell.value || '';
                                
                                try {
                                    // Handle different cell types
                                    if (cell.value && typeof cell.value === 'object') {
                                        if (cell.value.richText) {
                                            cellValue = cell.value.richText.map(rt => rt.text).join('');
                                        } else if (cell.value.formula) {
                                            cellValue = cell.value.result || cell.value.formula;
                                        } else if (cell.value.hyperlink) {
                                            cellValue = cell.value.text || cell.value.hyperlink;
                                        } else if (cell.value instanceof Date) {
                                            // Handle Date objects properly
                                            try {
                                                cellValue = cell.value.toLocaleDateString('nl-NL');
                                            } catch (e) {
                                                cellValue = 'Invalid Date';
                                            }
                                        } else {
                                            cellValue = cell.value.toString();
                                        }
                                    } else if (cell.value instanceof Date) {
                                        // Handle Date values that aren't wrapped in an object
                                        try {
                                            cellValue = cell.value.toLocaleDateString('nl-NL');
                                        } catch (e) {
                                            cellValue = 'Invalid Date';
                                        }
                                    }
                                } catch (cellError) {
                                    console.warn(`Error processing cell at row ${rowNumber}, col ${colNumber}:`, cellError);
                                    cellValue = 'Error';
                                }
                                
                                rowData[colNumber - 1] = cellValue;
                            });
                            jsonData.push(rowData);
                        });
                    } catch (processingError) {
                        console.error("Error processing Excel data:", processingError);
                        throw new Error(`Fout bij het verwerken van Excel-gegevens: ${processingError.message}`);
                    }
                    
                    if(jsonData.length === 0){
                        throw new Error("Het Excel-bestand is leeg of kon niet correct worden gelezen.");
                    }

                    setTableData(jsonData);
                } catch (e) {
                    console.error("Fout bij het verwerken van het Excel-bestand:", e);
                    setError(`Kon de gegevens niet laden. Details: ${e.message}`);
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

            const handleDrop = (event) => {
                event.preventDefault();
                const file = event.dataTransfer.files[0];
                if (file) {
                    if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                        file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                        processExcelFile(file);
                    } else {
                        setError('Selecteer een geldig Excel bestand (.xlsx of .xls)');
                    }
                }
            };

            const handleDragOver = (event) => {
                event.preventDefault();
            };
            
            // Effect voor het synchroniseren van de scrollbars
            useEffect(() => {
                const topScroll = topScrollRef.current;
                const tableContainer = tableContainerRef.current;
                
                if (!topScroll || !tableContainer) return;
                
                const topScrollContent = topScroll.querySelector('div');
                const dataTable = tableContainer.querySelector('.data-table');
                
                if (!topScrollContent || !dataTable) return;

                // Zet de breedte van de dummy content gelijk aan de tabelbreedte
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
                
                // Luister naar window resize om de breedte opnieuw in te stellen
                window.addEventListener('resize', setWidths);

                // Cleanup functie
                return () => {
                    topScroll.removeEventListener('scroll', handleTopScroll);
                    tableContainer.removeEventListener('scroll', handleTableScroll);
                    window.removeEventListener('resize', setWidths);
                };

            }, [loading]); // Opnieuw uitvoeren als loading verandert (wanneer de tabel verschijnt)


            return (
                <div className="page-wrapper">
                    <div className="container">
                        <header className="header">
                            <h1 className="title">Verkeersdingen Werklijst</h1>
                            <p className="description">
                                Upload uw Excel werklijst om verkeersborden gegevens te bekijken in een moderne, professionele interface. Sleep het bestand naar het upload gebied of klik om een bestand te selecteren.
                            </p>
                        </header>
                        
                        {!tableData.length && !loading && (
                            <div 
                                className="upload-area"
                                onDrop={handleDrop}
                                onDragOver={handleDragOver}
                                onClick={() => fileInputRef.current?.click()}
                            >
                                <div className="upload-content">
                                    <div className="upload-icon">📊</div>
                                    <h3>Sleep een .xlsx werklijst hier naartoe</h3>
                                    <p>of klik om een verkeersborden Excel bestand te selecteren</p>
                                    <button className="upload-button">Werklijst Selecteren</button>
                                </div>
                                <input
                                    ref={fileInputRef}
                                    type="file"
                                    accept=".xlsx,.xls"
                                    onChange={handleFileUpload}
                                    style={{ display: 'none' }}
                                />
                            </div>
                        )}

                        {fileName && (
                            <div className="file-info">
                                <span>Geladen bestand: <strong>{fileName}</strong></span>
                                <button 
                                    className="new-file-button"
                                    onClick={() => {
                                        setTableData([]);
                                        setFileName(null);
                                        setError(null);
                                        if (fileInputRef.current) {
                                            fileInputRef.current.value = '';
                                        }
                                    }}
                                >
                                    Nieuw Bestand
                                </button>
                            </div>
                        )}
                        
                        {loading && <div className="loading-indicator">Bestand verwerken...</div>}
                        {error && <div className="error-message">{error}</div>}

                        {!loading && !error && tableData.length > 0 && (
                            <React.Fragment>
                                <div className="top-scrollbar" ref={topScrollRef}>
                                    <div></div>
                                </div>
                                <section className="table-container" ref={tableContainerRef}>
                                    <table className="data-table">
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
                                                            {(() => {
                                                                // Safety check for rendering cell values
                                                                let displayValue = cell;
                                                                
                                                                // Handle null, undefined, or invalid values
                                                                if (cell === null || cell === undefined) {
                                                                    displayValue = '';
                                                                } else if (cell instanceof Date) {
                                                                    try {
                                                                        displayValue = cell.toLocaleDateString('nl-NL');
                                                                    } catch (e) {
                                                                        displayValue = 'Invalid Date';
                                                                    }
                                                                } else if (typeof cell === 'object') {
                                                                    displayValue = JSON.stringify(cell);
                                                                }
                                                                
                                                                // Check if it's a URL
                                                                if (typeof displayValue === "string" && displayValue.startsWith("http")) {
                                                                    return (
                                                                        <a
                                                                            href={displayValue}
                                                                            target="_blank"
                                                                            rel="noopener noreferrer"
                                                                            className="table-link"
                                                                        >
                                                                            Bekijk link
                                                                        </a>
                                                                    );
                                                                }
                                                                
                                                                return displayValue;
                                                            })()}
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

        ReactDOM.render(
            <ErrorBoundary>
                <VerkeersbordenWerklijst />
            </ErrorBoundary>, 
            document.getElementById("root")
        );
    </script>
</body>
</html>
