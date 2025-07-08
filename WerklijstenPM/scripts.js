// Haal de basislink op
var basislink = _spPageContextInfo.webAbsoluteUrl;

// Construeer het volledige bestandspad
var filePath = basislink + "/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten PM/Werklijsten MAPS PM Verkeersborden.xlsx";

// Helper function to convert ExcelJS worksheet to HTML table
function convertWorksheetToHTML(worksheet) {
  var html = '<table border="1" style="border-collapse: collapse;">';
  
  worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    html += '<tr>';
    row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
      var cellValue = cell.value || '';
      
      // Handle different cell types
      if (cell.value && typeof cell.value === 'object') {
        if (cell.value.richText) {
          cellValue = cell.value.richText.map(rt => rt.text).join('');
        } else if (cell.value.formula) {
          cellValue = cell.value.result || cell.value.formula;
        } else if (cell.value.hyperlink) {
          cellValue = cell.value.text || cell.value.hyperlink;
        } else {
          cellValue = cell.value.toString();
        }
      }
      
      html += '<td>' + cellValue + '</td>';
    });
    html += '</tr>';
  });
  
  html += '</table>';
  return html;
}
z
// Gebruik CallPost.js om de RequestDigestHeader te verkrijgen
CallPost.js({
  url: _spPageContextInfo.webAbsoluteUrl + "/_api/contextinfo",
  method: "POST",
  headers: {
    "Accept": "application/json; odata=verbose"
  },
  success: function (data) {
    var digest = data.d.GetContextWebInformation.FormDigestValue;

    // Fetch het Excelbestand
    fetch(filePath, {
      headers: {
        "Accept": "application/json; odata=verbose",
        "X-RequestDigest": digest
      }
    })
    .then(response => response.arrayBuffer())
    .then(data => {
      // Create new ExcelJS workbook and load the data
      var workbook = new ExcelJS.Workbook();
      return workbook.xlsx.load(data);
    })
    .then(workbook => {
      // Get the first worksheet
      var worksheet = workbook.worksheets[0];
      
      // Convert worksheet to HTML table
      var htmlstr = convertWorksheetToHTML(worksheet);
      document.getElementById('excelData').innerHTML = htmlstr;
    })
    .catch(error => {
      console.error("Fout bij het ophalen van het Excelbestand:", error);
      document.getElementById('excelData').innerHTML = "Fout bij het laden van de data.";
    });
  },
  error: function (error) {
    console.error("Fout bij het ophalen van RequestDigestHeader:", error);
    document.getElementById('excelData').innerHTML = "Fout bij het laden van de data.";
  }
});