// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const express = require('express');
const https = require('https');
const morgan = require('morgan');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');
const xml2js = require('xml2js');

const DEFAULT_PORT = process.env.PORT || 3000;

// Initialize express.
const app = express();

// Configure morgan module to log all requests.
app.use(morgan('dev'));

// Parse JSON bodies.
app.use(express.json());

// Set up app folders.
app.use(express.static('WebApplication/App'));

// Serve add-in files from root directory
app.use(express.static(__dirname));

// Serve manifest file
app.get('/manifest.xml', (req, res) => {
    res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// Serve taskpane files
app.get('/src/taskpane.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'src', 'taskpane.html'));
});

app.get('/src/taskpane.css', (req, res) => {
    res.sendFile(path.join(__dirname, 'src', 'taskpane.css'));
});

app.get('/src/taskpane.js', (req, res) => {
    res.sendFile(path.join(__dirname, 'src', 'taskpane.js'));
});

app.get('/src/commands.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'src', 'commands.html'));
});

// Serve icon files
app.use('/assets', express.static(path.join(__dirname, 'assets')));

// API endpoint to create spreadsheet.
// Security note: This API is public and can be called by any client. Be sure to add authentication and authorization for this API in a production environment.
app.post('/api/create-spreadsheet', async (req, res) => {
    try {
        const tableData = req.body;

        // Basic validation of request body structure.
        if (!tableData || !Array.isArray(tableData.rows)) {
            return res.status(400).json({ error: 'Invalid request body: "rows" must be an array.' });
        }
        
        // Create a new workbook and worksheet.
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Web Data');
        
        // Insert data from the request.
        tableData.rows.forEach((row, rowIndex) => {
            if (!row || !Array.isArray(row.columns)) {
                // Skip malformed rows instead of throwing.
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
        
        // Auto-fit columns.
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
        
        // Apply formatting to header row (first row).
        const headerRow = worksheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' }
        };
        
        // Embed the custom add-in and get the modified buffer.
        const buffer = await embedAddin(workbook);
        
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
 * Embeds the custom Office add-in defined in manifest.xml into the workbook.
 * This manipulates the Open Office XML (OOXML) structure to embed a web extension.
 * 
 * The process:
 * 1. Generates the workbook using ExcelJS.
 * 2. Loads the .xlsx file (which is a ZIP archive) using JSZip.
 * 3. Adds webextension XML files to configure the add-in.
 * 4. Updates [Content_Types].xml to register new parts.
 * 5. Updates workbook.xml.rels to link the taskpane.
 * 6. Re-packages the modified ZIP as an .xlsx file.
 * 
 * To embed your own add-in:
 * - Modify createWebExtensionXml() to change the add-in reference (ID, store, storeType).
 * - Modify createTaskpaneXml() to change visibility and other taskpane properties.
 * 
 * @param {ExcelJS.Workbook} workbook - The workbook to embed the add-in into
 * @returns {Promise<Buffer>} The modified workbook with embedded add-in
 */
async function embedAddin(workbook) {
    // First, generate the workbook as a buffer.
    const buffer = await workbook.xlsx.writeBuffer();
    
    // Load the buffer into JSZip.
    const zip = await JSZip.loadAsync(buffer);
    
    // Create the webextension part XML.
    const webExtensionXml = createWebExtensionXml();
    
    // Create the taskpane part XML.
    const taskpaneXml = createTaskpaneXml();
    
    // Add the webextension files to the zip.
    // JSZip automatically creates parent folders when adding files with paths.
    zip.file('xl/webextensions/webextension1.xml', webExtensionXml);
    zip.file('xl/webextensions/_rels/taskpanes.xml.rels', createWebExtensionRels());
    zip.file('xl/webextensions/taskpanes.xml', taskpaneXml);
    
    // Update or create [Content_Types].xml.
    await updateContentTypes(zip);
    
    // Update workbook.xml.rels.
    await updateWorkbookRels(zip);
    
    // Return the modified zip as a buffer.
    const modifiedBuffer = await zip.generateAsync({ 
        type: 'nodebuffer',
        compression: 'DEFLATE',
        compressionOptions: {
            level: 9
        }
    });
    
    return modifiedBuffer;
}

/**
 * Creates the webextension XML content for our custom add-in.
 * Note: The store path should match where the manifest.xml is sideloaded from.
 * For network share sideloading, use UNC path format: //COMPUTERNAME/ShareName/manifest.xml
 */
function createWebExtensionXml() {
    // Use "developer" for sideloaded add-ins as per Microsoft documentation
    // https://learn.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{12345678-1234-1234-1234-123456789012}">
    <we:reference id="12345678-1234-1234-1234-123456789012" version="1.0.0.0" store="developer" storeType="Registry"/>
    <we:alternateReferences/>
    <we:properties>
        <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
    </we:properties>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>`;
}

/**
 * Creates the taskpane XML content.
 * visibility="1" means the task pane will open automatically when the document is opened.
 * After opening manually once, Office.js code in the taskpane can control the auto-open setting.
 */
function createTaskpaneXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
    <wetp:taskpane dockstate="right" visibility="1" width="350" row="4">
        <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
    </wetp:taskpane>
</wetp:taskpanes>`;
}

/**
 * Creates the relationship file for taskpanes to reference the webextension.
 */
function createWebExtensionRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>`;
}

/**
 * Updates [Content_Types].xml to include webextension types.
 */
async function updateContentTypes(zip) {
    const contentTypesPath = '[Content_Types].xml';
    const contentTypesFile = zip.file(contentTypesPath);
    if (!contentTypesFile) {
        throw new Error(`Missing required file "${contentTypesPath}" in workbook package.`);
    }
    const contentTypesXml = await contentTypesFile.async('string');
    
    const parser = new xml2js.Parser();
    // Use compact XML (pretty: false) to keep the Office Open XML package small and consistent
    // with the original workbook. This makes manual debugging a bit harder but preserves the
    // existing file layout produced by Excel.
    const builder = new xml2js.Builder({ headless: false, renderOpts: { pretty: false } });
    
    const result = await parser.parseStringPromise(contentTypesXml);
    
    // Check if webextension types already exist.
    const defaults = result.Types.Default || [];
    const overrides = result.Types.Override || [];
    
    // Add webextension override if it doesn't exist.
    const webExtensionExists = overrides.some(o => 
        o.$.PartName === '/xl/webextensions/webextension1.xml'
    );
    
    if (!webExtensionExists) {
        overrides.push({
            $: {
                PartName: '/xl/webextensions/webextension1.xml',
                ContentType: 'application/vnd.ms-office.webextension+xml'
            }
        });
    }
    
    // Add taskpanes override if it doesn't exist.
    const taskpanesExists = overrides.some(o => 
        o.$.PartName === '/xl/webextensions/taskpanes.xml'
    );
    
    if (!taskpanesExists) {
        overrides.push({
            $: {
                PartName: '/xl/webextensions/taskpanes.xml',
                ContentType: 'application/vnd.ms-office.webextensiontaskpanes+xml'
            }
        });
    }
    
    result.Types.Override = overrides;
    
    const updatedXml = builder.buildObject(result);
    zip.file(contentTypesPath, updatedXml);
}

/**
 * Updates workbook.xml.rels to include reference to taskpanes.
 */
async function updateWorkbookRels(zip) {
    const relsPath = 'xl/_rels/workbook.xml.rels';
    const relsFile = zip.file(relsPath);
    if (!relsFile) {
        throw new Error(`Relationships file not found in workbook: ${relsPath}`);
    }
    const relsXml = await relsFile.async('string');
    
    const parser = new xml2js.Parser();
    const builder = new xml2js.Builder({ headless: false, renderOpts: { pretty: false } });
    
    const result = await parser.parseStringPromise(relsXml);
    
    // Find the highest rId.
    const relationships = result.Relationships.Relationship || [];
    let maxId = 0;
    relationships.forEach(rel => {
        const id = parseInt(rel.$.Id.replace('rId', ''));
        if (id > maxId) maxId = id;
    });
    
    // Check if taskpanes relationship already exists.
    const taskpanesExists = relationships.some(rel => 
        rel.$.Type === 'http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes'
    );
    
    if (!taskpanesExists) {
        relationships.push({
            $: {
                Id: `rId${maxId + 1}`,
                Type: 'http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes',
                Target: 'webextensions/taskpanes.xml'
            }
        });
        
        result.Relationships.Relationship = relationships;
        
        const updatedXml = builder.buildObject(result);
        zip.file(relsPath, updatedXml);
    }
}

app.get('/redirect', (req, res) => {
    res.sendFile(path.join(__dirname, 'WebApplication/App/redirect.html'));
});

// Set up a route for index.html.
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'WebApplication/App/index.html'));
});

// Start HTTPS server
const certPath = path.join(__dirname, 'localhost.crt');
const keyPath = path.join(__dirname, 'localhost.key');

if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    const httpsOptions = {
        key: fs.readFileSync(keyPath),
        cert: fs.readFileSync(certPath)
    };
    
    https.createServer(httpsOptions, app).listen(DEFAULT_PORT, () => {
        console.log(`Sample app listening on https://localhost:${DEFAULT_PORT}!`);
    });
} else {
    console.error('ERROR: SSL certificate files not found!');
    console.error(`Please ensure localhost.crt and localhost.key exist in the root folder.`);
    console.error(`Run 'npm run generate-cert' to generate self-signed certificates.`);
    process.exit(1);
}

module.exports = app;
