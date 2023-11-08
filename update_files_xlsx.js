const fs = require("fs");
const XLSX = require("xlsx");
const axios = require("axios");

async function appendDataToSecondSheet(sheet, payload) {
  const rowValues = [new Date().toLocaleString()]; // Add a timestamp as the first value

  // Append values from the payload to the row
  for (const key of Object.keys(payload)) {
    rowValues.push(payload[key]);
  }

  sheet.push(rowValues);
}

async function readExcelFilesInFolderAndSendRequests(
  folderPath,
  serviceEndpoint
) {
  const excelFiles = fs
    .readdirSync(folderPath)
    .filter((file) => file.endsWith(".xlsm"));

  for (const file of excelFiles) {
    const filePath = `${folderPath}/${file}`;

    const workbook = XLSX.readFile(filePath);

    const configsheet = workbook.Sheets["Configuration"]; // Change to your actual sheet name
    const datasheet = workbook.Sheets["Data"]; // Change to your actual sheet name

    const url = configsheet["B8"].v;

    const payload = {};
    let rowNumber = 11; // Start reading from row 10

    while (configsheet[`A${rowNumber}`]) {
      const cellA = configsheet[`A${rowNumber}`].v;
      const cellB = configsheet[`B${rowNumber}`].v;
      payload[cellA] = cellB;
      rowNumber++;
    }

    console.log(`Data from ${file}:`, payload);

    // Send the payload to the service endpoint
    try {
      const readyPayload = { url: url, cssSelectors: { ...payload } };
      const response = await axios.post(serviceEndpoint, readyPayload);
      console.log(`Service response for ${file}:`, response.data);

      // Append the response data to the second sheet
      //appendDataToSecondSheet(datasheet, response.data);

      const rowValues = [new Date().toLocaleString()]; // Add a timestamp as the first value
      // Append values from the response.data to the row
      for (const key of Object.keys(response.data)) {
        rowValues.push(response.data[key]);
      }
      //datasheet.push(rowValues);
      const wsData = [rowValues];
      XLSX.utils.sheet_add_aoa(datasheet, wsData, { origin: -1 });

      // Save the updated Excel file
      XLSX.writeFile(workbook, filePath);

      console.log(`Data appended to the second sheet of ${file}`);
    } catch (error) {
      console.error(`Error sending data to the service for ${file}:`, error);
    }
  }
}

//const folderPath = __dirname + "\\Databases"; // Replace with the path to your folder
//const serviceEndpoint = "http://localhost:3000/scrape"; // Replace with your service endpoint

//readExcelFilesInFolderAndSendRequests(folderPath, serviceEndpoint);
module.exports = { readExcelFilesInFolderAndSendRequests }