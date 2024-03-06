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
    console.log('Loading worksheet data...')
    sheet.load('name'); // Explicitly load the 'name' property

    const range = sheet.getUsedRange();
    range.load('values'); // Adjust as needed to load relevant data

    await context.sync();
    console.log('Serializing worksheet data...')
    // Serialize current worksheet data
    const worksheetData = JSON.stringify({
      sheetName: sheet.name, // Now you can safely access the 'name' property
      data: range.values
    });
    console.log(`Sending worksheet data: \n ${worksheetData}`)
    // Send serialized data to your API
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

