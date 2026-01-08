// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const express = require('express');
const morgan = require('morgan');
const path = require('path');
const ExcelJS = require('exceljs');

const DEFAULT_PORT = process.env.PORT || 3000;

// initialize express.
const app = express();

// Configure morgan module to log all requests.
app.use(morgan('dev'));

// Parse JSON bodies
app.use(express.json());

// Setup app folders.
app.use(express.static('WebApplication/App'));

// API endpoint to create spreadsheet
// Security note: This API is public and can be called by any client. Be sure to add authentication and authorization for this API in a production environment.
app.post('/api/create-spreadsheet', async (req, res) => {
    try {
        const tableData = req.body;

        // Basic validation of request body structure
        if (!tableData || !Array.isArray(tableData.rows)) {
            return res.status(400).json({ error: 'Invalid request body: "rows" must be an array.' });
        }
        
        // Create a new workbook and worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Web Data');
        
        // Insert data from the request
        tableData.rows.forEach((row, rowIndex) => {
            if (!row || !Array.isArray(row.columns)) {
                // Skip malformed rows instead of throwing
                return;
            }
            const excelRow = worksheet.getRow(rowIndex + 1);
            row.columns.forEach((column, colIndex) => {
                if (column && Object.prototype.hasOwnProperty.call(column, 'value')) {
                    excelRow.getCell(colIndex + 1).value = column.value;
                }
            });
            excelRow.commit();
        });
        
        // Auto-fit columns
        worksheet.columns.forEach((column, index) => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
                const columnLength = cell.value ? cell.value.toString().length : 10;
                if (columnLength > maxLength) {
                    maxLength = columnLength;
                }
            });
            column.width = maxLength < 10 ? 10 : maxLength + 2;
        });
        
        // Apply formatting to header row (first row)
        const headerRow = worksheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' }
        };
        
        // Embed Script Lab add-in
        await embedAddin(workbook);
        
        // Generate buffer and send as response
        const buffer = await workbook.xlsx.writeBuffer();
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=spreadsheet.xlsx');
        res.send(buffer);
        
    } catch (error) {
        console.error('Error creating spreadsheet:', error);
        const responseBody = { error: 'Failed to create spreadsheet' };
        if (process.env.NODE_ENV !== 'production' && error && error.message) {
            responseBody.details = error.message;
        }
        res.status(500).json(responseBody);
    }
});

/**
 * Embeds the Script Lab add-in into the workbook.
 * This uses the Open Office XML structure to embed a web extension.
 * @param {ExcelJS.Workbook} workbook - The workbook to embed the add-in into
 */
async function embedAddin(workbook) {
    // ExcelJS doesn't directly support web extensions, but we can add them
    // through the workbook's model after it's created
    // This is a simplified version - for full add-in embedding, 
    // you would need to manipulate the OOXML directly
    
    // Add a custom property to auto-show the task pane
    workbook.properties = workbook.properties || {};
    workbook.properties.custom = workbook.properties.custom || {};
    
    // Note: Full add-in embedding requires manipulating the OOXML zip structure
    // which ExcelJS doesn't directly support for web extensions.
    // For a complete implementation, you would need to:
    // 1. Generate the workbook
    // 2. Unzip it
    // 3. Add webextension.xml and webextensions.xml files
    // 4. Update [Content_Types].xml and _rels/.rels
    // 5. Re-zip the file
    
    // For this sample, we're providing the core functionality.
    // The workbook will be created and populated with data.
    // Users can manually add the add-in after opening the file if needed.
}

app.get('/redirect', (req, res) => {
    res.sendFile(path.join(__dirname, 'WebApplication/App/redirect.html'));
});

// Set up a route for index.html
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'WebApplication/App/index.html'));
});

app.listen(DEFAULT_PORT, () => {
    console.log(`Sample app listening on port ${DEFAULT_PORT}!`);
});

module.exports = app;
