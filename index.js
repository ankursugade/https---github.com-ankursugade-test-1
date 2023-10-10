
const express = require('express');
const app = express();
const port = process.env.PORT || 3000;
const sqlite3 = require('sqlite3').verbose();
const bodyParser = require('body-parser');

// Serve static files from the 'public' folder
app.use(express.static('public'));

// Set up SQLite database
const db = new sqlite3.Database('data.db');
db.serialize(() => {
    db.run('CREATE TABLE IF NOT EXISTS data ('
        + 'observationBy TEXT, '
        + 'observationDate TEXT, '
        + 'floor TEXT, '
        + 'area TEXT, '
        + 'observationDescription TEXT, '
        + 'observationType TEXT)');
});

// Parse incoming request data
app.use(bodyParser.urlencoded({ extended: false }));

// Define a route for the form
app.get('/', (req, res) => {
    res.sendFile(__dirname + '/views/form.html');
});

// Define a route to handle form submissions
app.post('/submit', (req, res) => {
    const {
        observationBy,
        observationDate,
        floor,
        area,
        observationDescription,
        observationType
    } = req.body;

    console.log('Received form data:');
    console.log('Observation By:', observationBy);
    console.log('Observation Date:', observationDate);
    console.log('Floor:', floor);
    console.log('Area:', area);
    console.log('Observation Description:', observationDescription);
    console.log('Observation Type:', observationType);

    // Validate the form data (basic validation)
    if (!observationBy || !observationDate || !floor || !area || !observationDescription) {
        console.error('Validation error: Required fields are missing.');
        return res.status(400).send('Error: Please fill out all required fields.');
    }

    // Insert the form data into the database
    db.run(
        'INSERT INTO data (observationBy, observationDate, floor, area, observationDescription, observationType) VALUES (?, ?, ?, ?, ?, ?)',
        [observationBy, observationDate, floor, area, observationDescription, observationType],
        (err) => {
            if (err) {
                console.error('Error inserting data:', err);
                return res.status(500).send('Error: Unable to submit the form.');
            }
            
            console.log('Data inserted successfully.');
            res.send('Form submitted successfully!');

            // Reset the form fields after successful submission
            resetForm();
        }
    );
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});

// --------
const XLSX = require('xlsx');

// Define a route to generate an Excel report
app.get('/generate-report', (req, res) => {
    // Get the start and end date from the query parameters (e.g., /generate-report?startDate=2023-01-01&endDate=2023-12-31)
    const startDate = req.query.startDate;
    const endDate = req.query.endDate;

    console.log('Generating Excel report for dates:', startDate, 'to', endDate);

    // Query the database to fetch data within the specified date range
    db.all(
        'SELECT * FROM data WHERE observationDate BETWEEN ? AND ?',
        [startDate, endDate],
        (err, data) => {
            if (err) {
                console.error('Error generating report:', err);
                return res.status(500).send('Error: Unable to generate the report.');
            }

            console.log('Generated report data:', data);

            // Create a new Excel workbook and worksheet
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet(data);

            // Add the worksheet to the workbook
            XLSX.utils.book_append_sheet(wb, ws, 'ObservationData');

            // Create a buffer containing the Excel file
            const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });

            // Set the response headers to indicate that it's an Excel file
            res.setHeader('Content-Disposition', 'attachment; filename=report.xlsx');
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.send(buffer);
        }
    );
});

// --- 
// Define a route to render the report generation page
app.get('/generate-report', (req, res) => {
    res.sendFile(__dirname + '/views/report.html');
});

// Modify the existing route for generating Excel reports (keep it as it is)
app.get('/generate-excel-report', (req, res) => {
    // ... (rest of the code for generating Excel reports)
});

// --- 
// Define a route to handle form submissions for generating Excel reports with date filters
app.post('/generate-excel-report', (req, res) => {
    const startDate = req.body.startDate;
    const endDate = req.body.endDate;

    console.log('Generating filtered Excel report for dates:', startDate, 'to', endDate);

    // Query the database to fetch data within the specified date range
    db.all(
        'SELECT * FROM data WHERE observationDate BETWEEN ? AND ?',
        [startDate, endDate],
        (err, data) => {
            if (err) {
                console.error('Error generating filtered report:', err);
                return res.status(500).send('Error: Unable to generate the report.');
            }

            console.log('Generated filtered report data:', data);

            // Create a new Excel workbook and worksheet
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet(data);

            // Add the worksheet to the workbook
            XLSX.utils.book_append_sheet(wb, ws, 'FilteredObservationData');

            // Create a buffer containing the Excel file
            const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });

            // Set the response headers to indicate that it's an Excel file
            res.setHeader('Content-Disposition', 'attachment; filename=filtered_report.xlsx');
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.send(buffer);
        }
    );
});

// Reset form fields
function resetForm() {
    document.getElementById('data-form').reset();
}

