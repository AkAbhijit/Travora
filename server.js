const express = require('express');
const fs = require('fs');
const path = require('path');
const bodyParser = require('body-parser');
const XLSX = require('xlsx'); // Import xlsx package

const app = express();
const PORT = 8080;

app.use(express.static(path.join(__dirname, 'public')));

// Parse form data
app.use(bodyParser.urlencoded({ extended: true }));

// Serve the tour.html as the default page (root URL)
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Serve the contact form page (index.html)
app.get('/contact', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'contact.html'));
});

// Handle form submission
app.post('/save-data', (req, res) => {
    const { name, email, phone, subject, reason, message } = req.body;

    // Data to be saved
    const formData = [
        { Name: name, Email: email, Phone: phone, Subject: subject, 'Reason for Contact': reason, Message: message }
    ];

    const filePath = 'form-submissions.xlsx';

    // Check if the Excel file exists
    if (fs.existsSync(filePath)) {
        // Read the existing workbook
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0]; // Assuming first sheet is where you want to save data
        const sheet = workbook.Sheets[sheetName];

        // Get the existing data and append the new form data
        const existingData = XLSX.utils.sheet_to_json(sheet);
        const updatedData = existingData.concat(formData);

        // Create a new sheet with updated data
        const updatedSheet = XLSX.utils.json_to_sheet(updatedData);
        workbook.Sheets[sheetName] = updatedSheet;

        // Write the updated data back to the file
        XLSX.writeFile(workbook, filePath);
    } else {
        // Create a new Excel file if it doesn't exist
        const workbook = XLSX.utils.book_new();
        const sheet = XLSX.utils.json_to_sheet(formData);

        // Append the sheet to the workbook
        XLSX.utils.book_append_sheet(workbook, sheet, 'Form Submissions');

        // Write the new Excel file
        XLSX.writeFile(workbook, filePath);
    }

    // Send success response with a cool design
    res.send(`
        <div style="text-align: center; padding: 50px; font-family: 'Arial', sans-serif;">
            <h1 style="color: #4CAF50; font-size: 2.5rem; font-weight: bold;">Thank You for Reaching Out!</h1>
            <p style="font-size: 1.2rem; color: #333; margin-top: 20px;">
                Your message has been successfully received. We'll get back to you as soon as possible. 😊
            </p>
            <p style="font-size: 1rem; color: #555; margin-top: 15px;">
                In the meantime, feel free to explore more amazing destinations!
            </p>
            <a href="/" style="display: inline-block; margin-top: 30px; background-color: #4CAF50; color: #fff; padding: 12px 30px; font-size: 1rem; text-decoration: none; border-radius: 5px; transition: background-color 0.3s ease;">
                Go Back to Homepage
            </a>
        </div>
    `);
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
