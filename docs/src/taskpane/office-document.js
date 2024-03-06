/* global Excel, console, fetch */

const insertText = async (text) => {
  try {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
      return context.sync();
    });
  } catch (error) {
    console.error("Error: " + error);
  }
};

const sendCurrentWorksheetData = () => {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Load the 'name' property of the worksheet
    sheet.load('name');

    // Get the used range of the current worksheet
    const range = sheet.getUsedRange();
    range.load('values'); // Load the 'values' property of the range

    // Sync the context to retrieve the loaded properties
    await context.sync();

    // Filter out completely empty rows and convert all values to strings
    const convertedValues = range.values
      .filter(row => row.some(cell => cell !== null && cell !== ''))
      .map(row => row.map(cell => cell ? cell.toString() : ''));

    // Serialize the current worksheet data
    const worksheetData = JSON.stringify({
      sheetName: sheet.name, // Accessing the 'name' property
      data: convertedValues
    });

    // Send the serialized data to your API
    fetch('https://intellisync.ai/process-excel', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: worksheetData,
    })
      .then(response => response.json())
      .then(data => console.log('Success:', data))
      .catch((error) => console.error('Error:', error));
  }).catch((error) => {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
};


export { insertText, sendCurrentWorksheetData };



