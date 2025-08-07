/*
 * Anki Tools - Office Add-in Commands
 * Excel integration for Anki flashcard exports
 */

/* global Office, Excel */

Office.onReady(() => {
  console.log("Anki Tools Add-in is ready");
});

// Backend API base URL
const API_BASE_URL = 'http://localhost:3001/api';

/**
 * Import Anki file into current worksheet
 * @param event {Office.AddinCommands.Event}
 */
async function importFromAnki(event) {
  try {
    await Excel.run(async (context) => {
      // Show file picker dialog
      Office.context.ui.displayDialogAsync(
        'https://localhost:3000/file-picker.html',
        { height: 60, width: 50 },
        async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const dialog = result.value;
            
            // Listen for message from file picker
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
              const message = JSON.parse(arg.message);
              
              if (message.type === 'fileSelected') {
                try {
                  // Send file to backend for processing
                  const formData = new FormData();
                  formData.append('file', message.file);
                  
                  const response = await fetch(`${API_BASE_URL}/import`, {
                    method: 'POST',
                    body: formData
                  });
                  
                  if (response.ok) {
                    const data = await response.json();
                    
                    // Check for filename issues and warn user
                    let warningMessage = '';
                    if (data.filenameValidation && (!data.filenameValidation.isValid || data.filenameValidation.warnings.length > 0)) {
                      warningMessage = '<div style="background:#fff3cd;padding:10px;margin:10px 0;border:1px solid #ffeaa7;">';
                      warningMessage += '<h4>⚠️ Filename Warning</h4>';
                      
                      if (!data.filenameValidation.isValid) {
                        warningMessage += '<p><strong>Issues found:</strong></p><ul>';
                        data.filenameValidation.issues.forEach(issue => {
                          warningMessage += `<li>${issue}</li>`;
                        });
                        warningMessage += '</ul>';
                        
                        if (data.filenameValidation.recommendation) {
                          warningMessage += `<p><strong>Recommended:</strong> ${data.filenameValidation.recommendation}</p>`;
                        }
                      }
                      
                      if (data.filenameValidation.warnings.length > 0) {
                        warningMessage += '<p><strong>Suggestions:</strong></p><ul>';
                        data.filenameValidation.warnings.forEach(warning => {
                          warningMessage += `<li>${warning}</li>`;
                        });
                        warningMessage += '</ul>';
                      }
                      
                      warningMessage += '<p><em>Consider renaming your Anki export file before importing for best results.</em></p>';
                      warningMessage += '</div>';
                    }
                    
                    // Clear existing worksheet
                    const worksheet = context.workbook.worksheets.getActiveWorksheet();
                    worksheet.getRange().clear();
                    
                    // Add headers and data
                    if (data.headers && data.rows) {
                      const headerRange = worksheet.getRange(`A1:${String.fromCharCode(64 + data.headers.length)}1`);
                      headerRange.values = [data.headers];
                      headerRange.format.font.bold = true;
                      
                      if (data.rows.length > 0) {
                        const dataRange = worksheet.getRange(`A2:${String.fromCharCode(64 + data.headers.length)}${data.rows.length + 1}`);
                        dataRange.values = data.rows;
                      }
                    }
                    
                    await context.sync();
                    
                    // Show success message with optional warnings
                    const successMessage = `
                      <html>
                        <head><style>body { font-family: Arial; padding: 20px; }</style></head>
                        <body>
                          <h3>✅ Import Successful</h3>
                          <p>Anki file imported successfully!</p>
                          ${warningMessage}
                        </body>
                      </html>
                    `;
                    
                    Office.context.ui.displayDialogAsync(
                      `data:text/html,${successMessage}`,
                      { height: warningMessage ? 60 : 30, width: 50 }
                    );
                  } else {
                    throw new Error('Import failed');
                  }
                } catch (error) {
                  console.error('Import error:', error);
                  Office.context.ui.displayDialogAsync(
                    `data:text/html,<html><body><h3>Import Error</h3><p>${error.message}</p></body></html>`,
                    { height: 30, width: 20 }
                  );
                }
              }
              
              dialog.close();
            });
          }
        }
      );
    });
  } catch (error) {
    console.error('Import error:', error);
    Office.context.ui.displayDialogAsync(
      `data:text/html,<html><body><h3>Import Error</h3><p>${error.message}</p></body></html>`,
      { height: 30, width: 20 }
    );
  }
  
  event.completed();
}

/**
 * Export current worksheet to Anki format
 * @param event {Office.AddinCommands.Event}
 */
async function exportToAnki(event) {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRange();
      
      if (usedRange) {
        usedRange.load("values");
        await context.sync();
        
        const response = await fetch(`${API_BASE_URL}/export`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            data: usedRange.values,
            sheetName: worksheet.name
          })
        });
        
        if (response.ok) {
          const result = await response.json();
          Office.context.ui.displayDialogAsync(
            `data:text/html,<html><body><h3>Export Successful</h3><p>File saved as: ${result.filename}</p></body></html>`,
            { height: 30, width: 20 }
          );
        } else {
          throw new Error('Export failed');
        }
      } else {
        throw new Error('No data to export');
      }
    });
  } catch (error) {
    console.error('Export error:', error);
    Office.context.ui.displayDialogAsync(
      `data:text/html,<html><body><h3>Export Error</h3><p>${error.message}</p></body></html>`,
      { height: 30, width: 20 }
    );
  }
  
  event.completed();
}

/**
 * Validate current worksheet format for Anki compatibility
 * @param event {Office.AddinCommands.Event}
 */
async function validateAnkiFormat(event) {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRange();
      
      if (usedRange) {
        usedRange.load("values");
        await context.sync();
        
        // Basic validation
        const values = usedRange.values;
        const headers = values[0];
        const isValid = headers.length >= 2 && values.length > 1;
        
        const message = isValid ? 
          `<h3>Validation Passed</h3><p>Worksheet has ${headers.length} columns and ${values.length - 1} data rows.</p>` :
          `<h3>Validation Failed</h3><p>Worksheet needs at least 2 columns and 1 data row.</p>`;
          
        Office.context.ui.displayDialogAsync(
          `data:text/html,<html><body>${message}</body></html>`,
          { height: 30, width: 20 }
        );
      } else {
        Office.context.ui.displayDialogAsync(
          'data:text/html,<html><body><h3>Validation Failed</h3><p>No data found in worksheet.</p></body></html>',
          { height: 30, width: 20 }
        );
      }
    });
  } catch (error) {
    console.error('Validation error:', error);
    Office.context.ui.displayDialogAsync(
      `data:text/html,<html><body><h3>Validation Error</h3><p>${error.message}</p></body></html>`,
      { height: 30, width: 20 }
    );
  }
  
  event.completed();
}

/**
 * Show help and instructions
 * @param event {Office.AddinCommands.Event}
 */
function showAnkiHelp(event) {
  const helpContent = `
    <html>
      <head><style>body { font-family: Arial; padding: 20px; }</style></head>
      <body>
        <h2>Anki Tools Help</h2>
        <h3>Usage:</h3>
        <ol>
          <li><strong>Import Anki:</strong> Import .txt file exported from Anki</li>
          <li><strong>Edit in Excel:</strong> Make changes to your flashcards</li>
          <li><strong>Export Anki:</strong> Save back to Anki-compatible format</li>
          <li><strong>Validate:</strong> Check format before export</li>
        </ol>
        <h3>Tips:</h3>
        <ul>
          <li>Keep column structure intact</li>
          <li>Use UTF-8 encoding for special characters</li>
          <li>Export files will have "-CLEANED" suffix</li>
        </ul>
      </body>
    </html>
  `;
  
  Office.context.ui.displayDialogAsync(
    `data:text/html,${helpContent}`,
    { height: 60, width: 50 }
  );
  
  event.completed();
}

// Register functions with Office
Office.actions.associate("importFromAnki", importFromAnki);
Office.actions.associate("exportToAnki", exportToAnki);
Office.actions.associate("validateAnkiFormat", validateAnkiFormat);
Office.actions.associate("showAnkiHelp", showAnkiHelp);
