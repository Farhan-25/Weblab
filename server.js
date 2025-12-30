const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3010;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve static files (HTML, CSS, JS)
app.use(express.static(__dirname));

// Excel file path
const excelFilePath = path.join(__dirname, 'data.xlsx');

// Function to read existing Excel file or create new one
function readOrCreateExcel() {
    if (fs.existsSync(excelFilePath)) {
        const workbook = XLSX.readFile(excelFilePath);
        return workbook;
    } else {
        // Create new workbook with headers
        const workbook = XLSX.utils.book_new();
        const headers = [['Name', 'Email', 'Company', 'Message', 'Date']];
        const worksheet = XLSX.utils.aoa_to_sheet(headers);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Contact Form Data');
        return workbook;
    }
}

// Function to append data to Excel file
function appendToExcel(data) {
    const workbook = readOrCreateExcel();
    const worksheet = workbook.Sheets['Contact Form Data'];
    
    // Convert sheet to JSON to get existing data
    const existingData = XLSX.utils.sheet_to_json(worksheet);
    
    // Add new row with timestamp
    const newRow = {
        'Name': data.name || '',
        'Email': data.email || '',
        'Company': data.company || '',
        'Message': data.message || '',
        'Date': new Date().toISOString()
    };
    
    existingData.push(newRow);
    
    // Convert back to worksheet
    const newWorksheet = XLSX.utils.json_to_sheet(existingData);
    workbook.Sheets['Contact Form Data'] = newWorksheet;
    
    // Write to file
    XLSX.writeFile(workbook, excelFilePath);
}

// API endpoint to handle form submission
app.post('/api/contact', (req, res) => {
    try {
        const { name, email, company, message } = req.body;
        
        // Validate required fields
        if (!name || !email || !message) {
            return res.status(400).json({ 
                success: false, 
                message: 'Name, email, and message are required fields.' 
            });
        }
        
        // Prepare data object
        const formData = {
            name: name.trim(),
            email: email.trim(),
            company: company ? company.trim() : '',
            message: message.trim()
        };
        
        // Append to Excel file
        appendToExcel(formData);
        
        // Send success response
        res.json({ 
            success: true, 
            message: 'Thank you for your message! We will get back to you soon.' 
        });
        
    } catch (error) {
        console.error('Error processing form submission:', error);
        res.status(500).json({ 
            success: false, 
            message: 'An error occurred while processing your request. Please try again later.' 
        });
    }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
    res.json({ status: 'OK', message: 'Server is running' });
});

// Start server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
    console.log(`Contact form data will be saved to: ${excelFilePath}`);
});
