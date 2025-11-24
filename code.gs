function openDateSelectorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Select Multiple Dates')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Function called by the client-side JavaScript to save the data
function saveDatesToCell(datesArray) {
  // Define variables locally at the start of the function
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeCell = sheet.getActiveCell();
  
  if (activeCell) {
    if (datesArray && datesArray.length > 0) {
      // Join the dates into a comma-separated string
      const dateString = datesArray.join(', '); 
      activeCell.setValue(dateString);
    } else {
      // Clear the cell if no dates are selected
      activeCell.clearContent();
    }
  }
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Create the standard top menu (this will appear immediately on open)
  ui.createMenu('📅 Show Dates') 
      .addItem('Multi-Date Picker', 'openDateSelectorSidebar') 
      .addToUi();
}

function getExistingDatesFromCell() {
  const cellValueRaw = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getValue();
  
  if (cellValueRaw) {
    const cellValueString = String(cellValueRaw).trim();
    
    // Split the string and format each part if it's a valid date
    const formattedDates = cellValueString.split(',').map(dateStr => {
      const dateObj = new Date(dateStr.trim());
      // Check if the date is valid and format it to MM/DD/YYYY
      if (dateObj instanceof Date && !isNaN(dateObj)) {
        const month = String(dateObj.getMonth() + 1).padStart(2, '0');
        const day = String(dateObj.getDate()).padStart(2, '0');
        const year = dateObj.getFullYear();
        return `${month}/${day}/${year}`;
      }
      return dateStr.trim(); // Return as is if not a valid date
    });
    
    return formattedDates;
  } else {
    return [];
  }
}


