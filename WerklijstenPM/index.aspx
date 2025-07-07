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
    <script src="https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/6.26.0/babel.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>

<body>
    <div id="root"></div>
    <script type="text/babel">
        const { useState, useEffect, useRef } = React;

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
                        const data = new Uint8Array(arrayBuffer);
                        const workbook = XLSX.read(data, { type: "array" });
                        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                        
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
