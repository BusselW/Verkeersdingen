<%@ Page Language="C#" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Verkeersborden Werklijst</title>
    <link href="styles.css" rel="stylesheet">
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
            const [loading, setLoading] = useState(true);
            const [error, setError] = useState(null);

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
                        console.error("Fout bij het ophalen of verwerken van het Excel-bestand:", e);
                        setError(`Kon de gegevens niet laden. Details: ${e.message}`);
                    } finally {
                        setLoading(false);
                    }
                };

                fetchExcelData();
            }, []);
            
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
                            <h1 className="title">Werklijst Verkeersborden</h1>
                            <p className="description">
                                Bekijk hieronder de gegevens uit het opgegeven Excel-bestand. Klik op het icoon om het Excelbestand te bewerken en de data op deze pagina te wijzigen.
                            </p>
                            <a
                                href="https://som.org.om.local/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten%20PM/Werklijsten%20MAPS%20PM%20Verkeersborden.xlsx?web=1"
                                className="download-icon"
                                target="_blank"
                                rel="noopener noreferrer"
                                title="Bewerk het Excel-bestand"
                            ></a>
                        </header>
                        
                        {loading && <div className="loading-indicator">Gegevens laden...</div>}
                        {error && <div className="error-message">{error}</div>}

                        {!loading && !error && (
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
