const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const bodyParser = require('body-parser');
const multer = require('multer');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Paths
const MASTER_FILE_PATH = path.join(__dirname, 'Master_Invoice_Records.xlsx');
const PRODUCTS_FILE_PATH = path.join(__dirname, 'data', 'products.json');
const LOGO_DIR = path.join(__dirname, 'public', 'images');

app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

// Ensure directories exist
if (!fs.existsSync(path.join(__dirname, 'data'))) fs.mkdirSync(path.join(__dirname, 'data'));
if (!fs.existsSync(LOGO_DIR)) fs.mkdirSync(LOGO_DIR, { recursive: true });

// --- LOGIN API ---
app.post('/api/login', (req, res) => {
    const { username, password } = req.body;
    // Default credentials: admin / admin123
    if (username === 'admin' && password === 'admin123') {
        res.json({ success: true });
    } else {
        res.status(401).json({ success: false });
    }
});

// --- PRODUCTS API ---
app.get('/api/products', (req, res) => {
    if (fs.existsSync(PRODUCTS_FILE_PATH)) {
        res.json(JSON.parse(fs.readFileSync(PRODUCTS_FILE_PATH)));
    } else {
        res.json([]);
    }
});

app.post('/api/products', (req, res) => {
    fs.writeFileSync(PRODUCTS_FILE_PATH, JSON.stringify(req.body, null, 2));
    res.json({ success: true });
});

// --- LOGO UPLOAD API ---
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, LOGO_DIR),
    filename: (req, file, cb) => cb(null, 'company-logo.jpg') // Always overwrite
});
const upload = multer({ storage: storage });

app.post('/api/upload-logo', upload.single('logo'), (req, res) => {
    res.json({ success: true });
});

// --- EXCEL LOGIC (SAVE INVOICE) ---
app.post('/api/save-invoice', async (req, res) => {
    const invoiceData = req.body;
    try {
        const workbook = new ExcelJS.Workbook();
        let allSheet, itemsSheet;

        if (fs.existsSync(MASTER_FILE_PATH)) {
            await workbook.xlsx.readFile(MASTER_FILE_PATH);
            allSheet = workbook.getWorksheet('All_Invoices');
            itemsSheet = workbook.getWorksheet('All_Items');
            
            if(!allSheet) allSheet = workbook.addWorksheet('All_Invoices');
            if(!itemsSheet) itemsSheet = workbook.addWorksheet('All_Items');
        } else {
            allSheet = workbook.addWorksheet('All_Invoices');
            itemsSheet = workbook.addWorksheet('All_Items');
        }

        // Set Headers (Ensures they exist)
        const summaryHeaders = [
            { header: 'Inv No', key: 'invNo', width: 15 },
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Buyer', key: 'buyer', width: 25 },
            { header: 'Total', key: 'total', width: 15 }
        ];
        const itemHeaders = [
            { header: 'Inv No', key: 'invNo', width: 15 },
            { header: 'Item', key: 'item', width: 25 },
            { header: 'Qty', key: 'qty', width: 10 },
            { header: 'Price', key: 'price', width: 15 },
            { header: 'Total', key: 'total', width: 15 }
        ];
        allSheet.columns = summaryHeaders;
        itemsSheet.columns = itemHeaders;

        // Add Summary Row
        allSheet.addRow({
            invNo: invoiceData.meta.invNo,
            date: invoiceData.meta.date,
            buyer: invoiceData.meta.to,
            total: invoiceData.totals.grandTotal
        });

        // Add Item Rows
        invoiceData.items.forEach(item => {
            itemsSheet.addRow({
                invNo: invoiceData.meta.invNo,
                item: item.description,
                qty: item.qty,
                price: item.price,
                total: item.rowTotal
            });
        });

        await workbook.xlsx.writeFile(MASTER_FILE_PATH);
        res.json({ success: true, message: 'Invoice Saved!' });

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: error.message });
    }
});

// --- EXCEL REPORT API ---
app.get('/api/reports', async (req, res) => {
    if (!fs.existsSync(MASTER_FILE_PATH)) return res.json({ invoices: [] });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(MASTER_FILE_PATH);
    const sheet = workbook.getWorksheet('All_Invoices');
    const invoices = [];
    if (sheet) {
        sheet.eachRow((row, i) => {
            if (i > 1) { // Skip header
                invoices.push({
                    invNo: row.getCell(1).value,
                    date: row.getCell(2).value,
                    buyer: row.getCell(3).value,
                    total: row.getCell(4).value
                });
            }
        });
    }
    res.json({ invoices });
});

app.listen(port, () => {
    console.log(`🚀 SpiceAdmin Pro running at http://localhost:${port}`);
});