// Haal de basislink op
var basislink = _spPageContextInfo.webAbsoluteUrl;

// Construeer het volledige bestandspad
var filePath = basislink + "/sites/MulderT/Onderdelen/Beoordelen/Verkeersborden/DocumentenVerkeersborden/Werklijsten PM/Werklijsten MAPS PM Verkeersborden.xlsx";

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
      var workbook = XLSX.read(data, {type: 'array'});

      var firstSheetName = workbook.SheetNames[0];
      var worksheet = workbook.Sheets[firstSheetName];

      var htmlstr = XLSX.utils.sheet_to_html(worksheet);
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