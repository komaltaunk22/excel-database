const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const app = express();
const PORT = 8080;
const RETRY_LIMIT = 5;
const RETRY_DELAY = 500; // milliseconds
var path = require('path');
app.use(bodyParser.json());
app.use(cors());
app.use(express.static("dist/superpowerbearing"));
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'dist/superpowerbearing/index.html'));
  });
  app.use(cors({
    methods: ['GET', 'POST'],
    origin: ["http://localhost:4200"],
 
}));
app.use(express.static('dist/assets/certificate'));
app.use(express.static('dist/assets/images'));
// Function to write data to the Excel file with retries
function writeFileWithRetry(workbook, filePath, retries = RETRY_LIMIT) {
    return new Promise((resolve, reject) => {
        function attemptWrite(attemptsLeft) {
            try {
                xlsx.writeFile(workbook, filePath);
                resolve();
            } catch (error) {
                if (attemptsLeft <= 0) {
                    reject(new Error('Failed to write file after multiple attempts'));
                } else {
                    console.log(`Retrying to write file, attempts left: ${attemptsLeft}`);
                    setTimeout(() => attemptWrite(attemptsLeft - 1), RETRY_DELAY);
                }
            }
        }
        attemptWrite(retries);
    });
}

// Read Excel file and convert to JSON
app.get('/api/viewdata', (req, res) => {
    const workbook = xlsx.readFile('data.xlsx');
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);
    res.json(data);
});

// Add data to Excel file
app.post('/api/data', async (req, res) => {
    try {
        const newData = req.body;
        const workbook = xlsx.readFile('data.xlsx');
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet);

        data.push(newData);
        const newSheet = xlsx.utils.json_to_sheet(data);
        workbook.Sheets[sheetName] = newSheet;

        await writeFileWithRetry(workbook, 'data.xlsx');

        res.json({ message: 'Data added successfully' });
        console.log('Super Power Bearing will contact you soon! Thank you');
    } catch (error) {
        console.error('Error writing to file:', error);
        res.status(500).json({ message: 'Try Again' });
    }
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
