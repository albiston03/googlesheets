function refreshImportXML() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets(); // Get all sheets in the spreadsheet

  // Loop through each sheet
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var range = sheet.getDataRange(); // Get the range of the entire sheet
    var formulas = range.getFormulas(); // Get all formulas in the range
    var displayValues = range.getDisplayValues(); // Get the displayed values (including errors)

    for (var i = 0; i < displayValues.length; i++) {
      for (var j = 0; j < displayValues[i].length; j++) {
        // Check if the cell displays '#N/A' and contains the 'IMPORTXML' function
        if (displayValues[i][j] === '#N/A' && formulas[i][j].toLowerCase().includes('importxml')) {
          var formula = formulas[i][j];
          
          // Extract the URL from the IMPORTXML formula
          var urlMatch = formula.match(/"(http[^"]+)"/);
          if (urlMatch) {
            var url = urlMatch[1];
            
            // Append a random query parameter to the URL to force a refresh
            var randomParam = Math.random().toString(36).substring(7); // Generate a random string
            var newUrl = url.includes('?') ? url + '&rand=' + randomParam : url + '?rand=' + randomParam;
            
            // Replace the original URL in the formula with the new one
            var newFormula = formula.replace(url, newUrl);
            
            // Re-enter the updated formula
            sheet.getRange(i + 1, j + 1).setFormula(newFormula);
          }
        }
      }
    }
  }
}
