const express = require('express');
const ExcelJS = require('exceljs');
const app = express();
const port = 3000;

// Middleware to parse JSON bodies
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Function to load the Excel file and extract the Lookup Data sheet
async function loadLookupData() {
    const workbook = new ExcelJS.Workbook();

    // Read the Excel file (replace with your actual path if needed)
    await workbook.xlsx.readFile('D:/MSME Calculator/MSME TAM Calculator.xlsx');

    // Get the 'Lookup Data' sheet
    const worksheet = workbook.getWorksheet('Lookup Data');
    const lookupData = [];

    // Iterate over rows and extract data
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {  // Skip header row
            lookupData.push({
                productType: row.getCell(3).value,  // Column C: Industry/Product Type
                totalCompanies: row.getCell(4).value  // Column D: Total companies
            });
        }
    });

    return lookupData;  // Return the extracted data
}

// Endpoint to calculate TAM based on user inputs
app.post('/calculate-tam', async (req, res) => {
    const { sectorProductType, price } = req.body;

    // Load the Lookup Data
    const lookupData = await loadLookupData();

    // Find the relevant entry from the Lookup Data
    const lookupEntry = lookupData.find(entry => entry.productType === sectorProductType);
    if (!lookupEntry) {
        return res.status(404).json({ error: 'Sector/Product Type not found' });
    }

    const numberOfCompanies = lookupEntry.totalCompanies;
    const tam = price * numberOfCompanies;  // Calculate TAM (Price * Number of Companies)

    // Return the calculated TAM to the frontend
    res.json({ tam });
});

// Start the server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
