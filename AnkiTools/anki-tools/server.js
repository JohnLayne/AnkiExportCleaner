/**
 * Anki Tools Backend Server
 * Node.js REST API to bridge Office Add-in with Python processing
 */

const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const https = require('https');
const { spawn } = require('child_process');
const multer = require('multer');

const app = express();
const PORT = 3001;

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));

// File upload configuration
const upload = multer({ dest: 'uploads/' });

// Get project root directory (two levels up from server.js)
const projectRoot = path.resolve(__dirname, '../../');
const pythonScript = path.join(projectRoot, 'anki_excel_tool.py');

/**
 * Validate filename for Anki compatibility
 * @param {string} filename - The filename to validate
 * @returns {object} - Validation result with warnings
 */
function validateAnkiFilename(filename) {
  const issues = [];
  const warnings = [];
  
  // Check length
  if (filename.length > 50) {
    issues.push(`Filename too long (${filename.length} chars). Recommended: under 50 characters.`);
  }
  
  // Check for problematic characters
  const problematicChars = /[\/\\:*?"<>|]/;
  if (problematicChars.test(filename)) {
    issues.push('Contains special characters (/, \\, :, *, ?, ", <, >, |) that may cause issues.');
  }
  
  // Check for double underscores (common in Anki exports)
  if (filename.includes('__')) {
    warnings.push('Contains double underscores (__) - consider using single underscores or hyphens.');
  }
  
  // Check for spaces
  if (filename.includes(' ')) {
    warnings.push('Contains spaces - consider using underscores or hyphens for better compatibility.');
  }
  
  return {
    isValid: issues.length === 0,
    issues,
    warnings,
    recommendation: issues.length > 0 ? generateFilenameRecommendation(filename) : null
  };
}

/**
 * Generate a recommended filename
 * @param {string} originalFilename - Original filename
 * @returns {string} - Recommended filename
 */
function generateFilenameRecommendation(originalFilename) {
  return originalFilename
    .replace(/[\/\\:*?"<>|]/g, '') // Remove special chars
    .replace(/\s+/g, '_') // Replace spaces with underscores
    .replace(/__+/g, '_') // Replace multiple underscores with single
    .replace(/[_-]+/g, '_') // Normalize separators
    .substring(0, 45) // Truncate to safe length
    + '.txt';
}

/**
 * Import Anki file endpoint
 * POST /api/import
 */
app.post('/api/import', upload.single('file'), async (req, res) => {
  try {
    console.log('Import request received');
    
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    const filename = req.file.originalname;
    const filenameValidation = validateAnkiFilename(filename);
    
    if (!filenameValidation.isValid) {
      console.warn('Problematic filename detected:', filename);
      console.warn('Issues:', filenameValidation.issues);
    }
    
    // Read the uploaded file content
    const fileContent = req.file.buffer.toString('utf8');
    console.log('File content length:', fileContent.length);
    
    // Parse the Anki export format (tab-separated values)
    const lines = fileContent.split('\n').filter(line => line.trim());
    
    if (lines.length === 0) {
      return res.status(400).json({ error: 'File appears to be empty' });
    }
    
    // Parse headers and data
    const headers = lines[0].split('\t');
    const rows = lines.slice(1).map(line => line.split('\t'));
    
    console.log('Parsed headers:', headers);
    console.log('Parsed rows:', rows.length);
    
    const result = {
      headers: headers,
      rows: rows,
      filename: filename,
      filenameValidation: filenameValidation
    };
    
    res.json(result);
    
  } catch (error) {
    console.error('Import error:', error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * Export to Anki format endpoint
 * POST /api/export
 */
app.post('/api/export', async (req, res) => {
  try {
    const { data, sheetName } = req.body;
    
    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ error: 'Invalid data format' });
    }
    
    console.log('Export request received for sheet:', sheetName);
    console.log('Data rows:', data.length);
    
    // Convert Excel data to tab-separated format
    const tsvContent = data.map(row => row.join('\t')).join('\n');
    
    // Generate filename with -CLEANED suffix
    const baseFileName = sheetName || 'AnkiExport';
    const cleanFileName = baseFileName.replace(/-EXCEL$/, '') + '-CLEANED.txt';
    const outputPath = path.join(projectRoot, 'samples', cleanFileName);
    
    // Write file with UTF-8 encoding
    fs.writeFileSync(outputPath, tsvContent, 'utf8');
    
    console.log('File exported to:', outputPath);
    
    res.json({ 
      success: true,
      filename: cleanFileName,
      path: outputPath
    });
    
  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * Process Anki file with Python script
 * POST /api/process
 */
app.post('/api/process', async (req, res) => {
  try {
    const { inputFile, outputFile } = req.body;
    
    const pythonProcess = spawn('python', [
      pythonScript,
      '--input', inputFile,
      '--output', outputFile
    ]);
    
    let output = '';
    let error = '';
    
    pythonProcess.stdout.on('data', (data) => {
      output += data.toString();
    });
    
    pythonProcess.stderr.on('data', (data) => {
      error += data.toString();
    });
    
    pythonProcess.on('close', (code) => {
      if (code === 0) {
        res.json({ success: true, output });
      } else {
        res.status(500).json({ error: error || 'Python script failed' });
      }
    });
    
  } catch (error) {
    console.error('Process error:', error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * Health check endpoint
 */
app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'OK',
    timestamp: new Date().toISOString(),
    pythonScript: fs.existsSync(pythonScript) ? 'Found' : 'Missing'
  });
});

/**
 * Get project info
 */
app.get('/api/info', (req, res) => {
  res.json({
    projectRoot,
    pythonScript,
    sampleFiles: fs.readdirSync(path.join(projectRoot, 'samples')).filter(f => f.endsWith('.txt'))
  });
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error('Server error:', error);
  res.status(500).json({ error: 'Internal server error' });
});

// Start server with HTTP (simpler for development)
app.listen(PORT, () => {
  console.log(`Anki Tools Backend Server running on http://localhost:${PORT}`);
  console.log(`Project root: ${projectRoot}`);
  console.log(`Python script: ${pythonScript}`);
  console.log('API endpoints:');
  console.log('  POST /api/import - Import Anki file');
  console.log('  POST /api/export - Export to Anki format');
  console.log('  POST /api/process - Process with Python script');
  console.log('  GET  /api/health - Health check');
  console.log('  GET  /api/info - Project information');
});

module.exports = app;
