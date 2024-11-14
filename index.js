const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const storage = multer.diskStorage({
  destination: 'uploads/',
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const fileFilter = (req, file, cb) => {
  if (file.mimetype.includes('spreadsheet') || 
      file.mimetype.includes('excel') ||
      file.originalname.endsWith('.xlsx') ||
      file.originalname.endsWith('.xls')) {
    cb(null, true);
  } else {
    cb(new Error('Invalid file type. Only Excel files are allowed.'), false);
  }
};

const upload = multer({ 
  storage: storage,
  fileFilter: fileFilter,
  limits: { fileSize: 5 * 1024 * 1024 }
});

const app = express();

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

const COLUMN_MAPPING = {
  'GSTIN of supplier': 'A',
  'Trade/Legal name': 'B',
  'Invoice number': 'C',
  'Invoice Date': 'E',
  'Invoice Value(₹)': 'F',
  'Taxable Value (₹)': 'I'
};

function formatCurrency(value) {
  if (typeof value === 'number') {
    return new Intl.NumberFormat('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(value);
  }
  return value;
}

function processExcelData(sheet) {
  const data = [];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  
  // Skip header rows (start from row 7 as per the Excel)
  for (let row = 6; row <= range.e.r; row++) {
    const rowData = {};
    
    // Check if row is not empty by checking GSTIN (column A)
    const gstinCell = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
    if (!gstinCell || !gstinCell.v) continue;

    Object.entries(COLUMN_MAPPING).forEach(([key, col]) => {
      const colIndex = XLSX.utils.decode_col(col);
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: colIndex })];
      rowData[key] = cell ? cell.v : '';
    });

    // Only add row if it has valid data
    if (Object.values(rowData).some(val => val)) {
      data.push(rowData);
    }
  }

  return data;
}

app.post('/process', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets['B2B'];
    
    if (!sheet) {
      return res.status(400).json({ error: 'B2B sheet not found in the Excel file' });
    }

    const data = processExcelData(sheet);

    if (data.length === 0) {
      return res.status(400).json({ error: 'No valid data found in the Excel file' });
    }

    const doc = new PDFDocument({ 
      layout: 'landscape',
      margin: 30
    });

    const pdfPath = path.join(__dirname, 'output.pdf');
    const writeStream = fs.createWriteStream(pdfPath);
    doc.pipe(writeStream);

    // Add title
    doc.font('Helvetica-Bold')
       .fontSize(16)
       .text('Goods and Services Tax - GSTR-2B', { align: 'center' });
    doc.moveDown(0.5);

    doc.fontSize(12)
       .text('Taxable inward supplies received from registered persons', { align: 'center' });
    doc.moveDown(1);

    // Table settings
    const tableTop = doc.y;
    const pageWidth = doc.page.width - 60;
    const columnWidths = {
      'GSTIN of supplier': pageWidth * 0.22,
      'Trade/Legal name': pageWidth * 0.23,
      'Invoice number': pageWidth * 0.15,
      'Invoice Date': pageWidth * 0.12,
      'Invoice Value(₹)': pageWidth * 0.14,
      'Taxable Value (₹)': pageWidth * 0.14
    };

    // Draw table header
    let xPos = 30;
    doc.font('Helvetica-Bold').fontSize(10);
    
    // Header background
    doc.fillColor('#f0f0f0')
       .rect(xPos, tableTop, pageWidth, 20)
       .fill();
    
    // Header text
    doc.fillColor('#000000');
    Object.entries(columnWidths).forEach(([header, width]) => {
      doc.text(
        header,
        xPos + 5,
        tableTop + 5,
        {
          width: width - 10,
          align: ['Invoice Value(₹)', 'Taxable Value (₹)'].includes(header) ? 'right' : 'left'
        }
      );
      xPos += width;
    });

    // Draw table rows
    let yPos = tableTop + 20;
    doc.font('Helvetica').fontSize(9);

    data.forEach((row, i) => {
      // Check if we need a new page
      if (yPos > doc.page.height - 50) {
        doc.addPage({ layout: 'landscape', margin: 30 });
        yPos = 50;
      }

      // Alternate row background
      if (i % 2 === 0) {
        doc.fillColor('#f9f9f9')
           .rect(30, yPos, pageWidth, 20)
           .fill();
      }

      // Row data
      doc.fillColor('#000000');
      xPos = 30;
      Object.entries(columnWidths).forEach(([header, width]) => {
        const value = row[header];
        const formattedValue = ['Invoice Value(₹)', 'Taxable Value (₹)'].includes(header)
          ? formatCurrency(value)
          : value;

        doc.text(
          formattedValue?.toString() || '',
          xPos + 5,
          yPos + 5,
          {
            width: width - 10,
            align: ['Invoice Value(₹)', 'Taxable Value (₹)'].includes(header) ? 'right' : 'left'
          }
        );
        xPos += width;
      });

      yPos += 20;
    });

    // Draw table borders
    doc.lineWidth(0.5);
    
    // Vertical lines
    xPos = 30;
    Object.values(columnWidths).forEach(width => {
      doc.moveTo(xPos, tableTop)
         .lineTo(xPos, yPos)
         .stroke();
      xPos += width;
    });
    doc.moveTo(xPos, tableTop)
       .lineTo(xPos, yPos)
       .stroke();

    // Horizontal lines
    for (let y = tableTop; y <= yPos; y += 20) {
      doc.moveTo(30, y)
         .lineTo(xPos, y)
         .stroke();
    }

    doc.end();

    writeStream.on('finish', () => {
      res.download(pdfPath, 'processed.pdf', (err) => {
        if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath);
      });
    });

  } catch (error) {
    console.error('Processing error:', error);
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    res.status(500).json({ error: error.message || 'Error processing file' });
  }
});

app.listen(3000, () => {
  console.log('Server running on port 3000');
});