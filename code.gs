/**
 * @OnlyCurrentDoc
 * This script adds a custom menu to the spreadsheet to process submitted links,
 * merges the data, removes duplicates, runs a report, and acts as a Web App.
 */

// --- CONFIGURATION ---
const SUBMISSION_SHEET = 'Submissions'; 
const SOURCE_DATA_SHEET_NAME = 'Sheet1'; 
const URL_COLUMN_NUMBER = 1;
// --- END CONFIGURATION ---

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SUBMISSION_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(SUBMISSION_SHEET);
      sheet.appendRow(['Sheet Link', 'Sheet Pin', 'Payment Method', 'Payment Number', 'User ID', 'Timestamp', 'Submission Type']);
    }
    const data = JSON.parse(e.postData.contents);
    sheet.appendRow([
      data.sheetLink, data.sheetPin, data.paymentMethod, 
      data.paymentNumber, data.userId, data.timestamp, data.submissionType
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ 'status': 'success', 'message': 'Row added successfully' }))
      .setMimeType(ContentService.MimeType.JSON)
      .withSuccessCode(200);
  } catch (error) {
    Logger.log(error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ 'status': 'error', 'message': error.toString() }))
      .setMimeType(ContentService.MimeType.JSON)
      .withSuccessCode(500);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Automation Bot')
    .addItem('Merge All Submitted Sheets', 'runMergeProcess')
    .addItem('Run Report & Update Payment', 'runReportAndUpdatePayment')
    .addToUi();
}

function getCurrentDate() {
  const today = new Date();
  const day = String(today.getDate()).padStart(2, '0');
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const year = today.getFullYear();
  return `${day}-${month}-${year}`;
}

function getSubmissionTypeFromTitle(title) {
  if (title.includes('0 FD')) return '0 FD';
  if (title.includes('30 FD')) return '30 FD';
  return 'Unknown';
}

function runMergeProcess() {
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert(
    'Start Merge Process?',
    'This will create date-based sheets, merge data, remove duplicates, and create payment tracking. This may take a few minutes. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) {
    ui.alert('Process cancelled.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const submissionSheet = ss.getSheetByName(SUBMISSION_SHEET);
  
  if (!submissionSheet) {
    ui.alert('Error: The sheet named "' + SUBMISSION_SHEET + '" was not found.');
    return;
  }

  const lastRow = submissionSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No submitted links to process.');
    return;
  }

  const currentDate = getCurrentDate();
  const urlRange = submissionSheet.getRange(2, URL_COLUMN_NUMBER, lastRow - 1, 7);
  const submissions = urlRange.getValues().filter(row => row[0]); // Filter out empty rows

  // Group submissions by type
  const submissionsByType = {};
  submissions.forEach(row => {
    const submissionType = getSubmissionTypeFromTitle(row[6]); // Submission Type column
    if (!submissionsByType[submissionType]) {
      submissionsByType[submissionType] = [];
    }
    submissionsByType[submissionType].push(row);
  });

  let totalProcessed = 0;
  let totalErrors = 0;
  let paymentData = [];

  ss.toast('Starting merge process... Please wait.', 'Processing', -1);

  // Process each submission type
  for (const [type, typeSubmissions] of Object.entries(submissionsByType)) {
    const sheetName = `${currentDate} ${type}`;
    let mergedSheet = ss.getSheetByName(sheetName);
    if (!mergedSheet) {
      mergedSheet = ss.insertSheet(sheetName);
    }
    
    mergedSheet.clear();
    SpreadsheetApp.flush();

    let allData = [];
    let masterHeader = null;
    let headerWritten = false;
    let errorLog = [];

    // Process each submission for this type
    for (const submission of typeSubmissions) {
      try {
        const url = submission[0];
        const paymentMethod = submission[2];
        const paymentNumber = submission[3];
        const userId = submission[4];

        const idMatch = url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (!idMatch) { throw new Error("Invalid URL"); }
        const sourceId = idMatch[1];
        
        const file = DriveApp.getFileById(sourceId);
        const sourceSpreadsheet = SpreadsheetApp.open(file);
        const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_DATA_SHEET_NAME);

        if (sourceSheet) {
          const data = sourceSheet.getDataRange().getValues();
          if (data.length === 0) { continue; }

          if (!headerWritten) {
            masterHeader = data[0];
            allData.push(...data);
            headerWritten = true;
          } else {
            const isDuplicateHeader = JSON.stringify(data[0]) === JSON.stringify(masterHeader);
            if (isDuplicateHeader) {
              if (data.length > 1) { allData.push(...data.slice(1)); }
            } else {
              allData.push(...data);
            }
          }

          // Add to payment data
          const idCount = data.length > 1 ? data.length - 1 : 0; // Exclude header
          if (idCount > 0) {
            paymentData.push({
              method: paymentMethod,
              number: paymentNumber,
              totalId: idCount,
              amount: 0, // Will be calculated later
              userId: userId,
              type: type
            });
          }

          totalProcessed++;
        } else {
          errorLog.push(`Sheet '${SOURCE_DATA_SHEET_NAME}' not found in: ${url}`);
          totalErrors++;
        }
      } catch (e) {
        errorLog.push(`Could not access URL: ${submission[0]}. Error: ${e.message}.`);
        totalErrors++;
      }
    }

    // Write data to merged sheet
    if (allData.length > 0) {
      mergedSheet.getRange(1, 1, allData.length, allData[0].length).setValues(allData);
      
      // Remove duplicates
      const dataRange = mergedSheet.getDataRange();
      const values = dataRange.getValues();
      const uniqueRows = {};
      let duplicatesRemoved = 0;

      for (let i = values.length - 1; i >= 1; i--) {
        const rowString = values[i].join('||');
        if (uniqueRows[rowString]) {
          mergedSheet.deleteRow(i + 1);
          duplicatesRemoved++;
        } else {
          uniqueRows[rowString] = true;
        }
      }
    }
  }

  // Create payment tracking sheet
  createPaymentSheet(ss, currentDate, paymentData);

  ss.toast('Merge process complete.', 'Success!', 5);
  
  let successMessage = `✅ Success! Process complete.\n\n` +
                       `Total Submissions Processed: ${totalProcessed}\n` +
                       `Total Errors: ${totalErrors}\n` +
                       `Payment Entries Created: ${paymentData.length}`;
  
  if (totalErrors > 0) {
    successMessage += `\n\nSome errors occurred during processing.`;
  }
  
  ui.alert(successMessage);
}

function createPaymentSheet(ss, currentDate, paymentData) {
  const paymentSheetName = `${currentDate} Payment`;
  let paymentSheet = ss.getSheetByName(paymentSheetName);
  if (!paymentSheet) {
    paymentSheet = ss.insertSheet(paymentSheetName);
  }
  
  paymentSheet.clear();
  
  // Add headers
  paymentSheet.getRange(1, 1, 1, 5).setValues([['Method', 'Number', 'Total Id', 'Amount', 'User ID']]);
  
  // Add payment data
  if (paymentData.length > 0) {
    const paymentRows = paymentData.map(item => [
      item.method,
      item.number,
      item.totalId,
      item.amount,
      item.userId
    ]);
    
    paymentSheet.getRange(2, 1, paymentRows.length, 5).setValues(paymentRows);
  }
  
  // Format headers
  paymentSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
}

function runReportAndUpdatePayment() {
  const ui = SpreadsheetApp.getUi();
  
  // Ask for rate
  const ratePrompt = ui.prompt(
    'Enter Rate',
    'Please enter the rate per ID (e.g., 0.5 for 50 paisa per ID):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (ratePrompt.getSelectedButton() !== ui.Button.OK || !ratePrompt.getResponseText()) {
    ui.alert('Rate input cancelled.');
    return;
  }
  
  const rate = parseFloat(ratePrompt.getResponseText());
  if (isNaN(rate) || rate < 0) {
    ui.alert('Invalid rate. Please enter a valid number.');
    return;
  }
  
  // Ask for report file URL
  const reportPrompt = ui.prompt(
    'Report File URL',
    'Please enter the URL of the Report File to check the status of the merged data:',
    ui.ButtonSet.OK_CANCEL
  );

  if (reportPrompt.getSelectedButton() !== ui.Button.OK || !reportPrompt.getResponseText()) {
    ui.alert('Report URL input cancelled.');
    return;
  }

  const reportUrl = reportPrompt.getResponseText();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentDate = getCurrentDate();
  
  try {
    // Get report data
    const reportIdMatch = reportUrl.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!reportIdMatch) { throw new Error("Invalid Report URL"); }
    const reportId = reportIdMatch[1];
    
    const reportFile = DriveApp.getFileById(reportId);
    const reportSpreadsheet = SpreadsheetApp.open(reportFile);
    const reportSheet = reportSpreadsheet.getSheets()[0];
    
    // Get all data from report sheet and check for colored cells
    const reportDataRange = reportSheet.getDataRange();
    const reportValues = reportDataRange.getValues();
    const reportBackgrounds = reportDataRange.getBackgrounds();
    
    // Find good IDs (cells with background color)
    const goodIds = new Set();
    for (let i = 0; i < reportValues.length; i++) {
      for (let j = 0; j < reportValues[i].length; j++) {
        const cellValue = String(reportValues[i][j]).trim();
        const backgroundColor = reportBackgrounds[i][j];
        
        // Check if cell has background color (not white/transparent)
        if (cellValue && backgroundColor && backgroundColor !== '#ffffff' && backgroundColor !== '') {
          // Extract numeric part from the cell value
          const numericValue = cellValue.replace(/\D/g, '');
          if (numericValue) {
            goodIds.add(numericValue);
          }
        }
      }
    }
    
    // Process each merged sheet
    const sheetTypes = ['0 FD', '30 FD'];
    let totalBadIds = 0;
    let totalGoodIds = 0;
    
    for (const type of sheetTypes) {
      const sheetName = `${currentDate} ${type}`;
      const mergedSheet = ss.getSheetByName(sheetName);
      
      if (mergedSheet) {
        const dataRange = mergedSheet.getDataRange();
        const values = dataRange.getValues();
        const backgrounds = dataRange.getBackgrounds();
        
        let badCount = 0;
        let goodCount = 0;
        
        // Check each cell in the sheet
        for (let i = 0; i < values.length; i++) {
          for (let j = 0; j < values[i].length; j++) {
            const cellValue = String(values[i][j]).trim();
            const numericValue = cellValue.replace(/\D/g, '');
            
            if (numericValue) {
              if (goodIds.has(numericValue)) {
                goodCount++;
                // Color good cells green
                mergedSheet.getRange(i + 1, j + 1).setBackground('#90EE90');
              } else {
                badCount++;
                // Color bad cells red
                mergedSheet.getRange(i + 1, j + 1).setBackground('#FFB6C1');
              }
            }
          }
        }
        
        totalGoodIds += goodCount;
        totalBadIds += badCount;
        
        // Add status column
        mergedSheet.getRange(1, values[0].length + 1).setValue('Status');
        mergedSheet.getRange(1, values[0].length + 1).setFontWeight('bold');
        
        for (let i = 1; i < values.length; i++) {
          const rowHasGoodId = values[i].some(cell => {
            const numericValue = String(cell).replace(/\D/g, '');
            return numericValue && goodIds.has(numericValue);
          });
          
          const status = rowHasGoodId ? 'Good' : 'Bad';
          mergedSheet.getRange(i + 1, values[0].length + 1).setValue(status);
        }
      }
    }
    
    // Update payment sheet with calculated amounts and adjust for bad IDs
    updatePaymentSheetWithAmounts(ss, currentDate, rate, totalBadIds);
    
    ss.toast('Report processing complete.', 'Success!', 5);
    
    const successMessage = `✅ Report processing complete!\n\n` +
                          `Rate per ID: ${rate}\n` +
                          `Total Good IDs: ${totalGoodIds}\n` +
                          `Total Bad IDs: ${totalBadIds}\n` +
                          `Payment sheet updated with calculated amounts.`;
    
    ui.alert(successMessage);
    
  } catch (e) {
    ui.alert(`❌ Error processing report: ${e.message}`);
  }
}

function updatePaymentSheetWithAmounts(ss, currentDate, rate, totalBadIds) {
  const paymentSheetName = `${currentDate} Payment`;
  const paymentSheet = ss.getSheetByName(paymentSheetName);
  
  if (!paymentSheet) {
    ui.alert(`Payment sheet '${paymentSheetName}' not found.`);
    return;
  }
  
  const dataRange = paymentSheet.getDataRange();
  const values = dataRange.getValues();
  
  // Calculate amounts for each payment entry
  for (let i = 1; i < values.length; i++) {
    const totalId = parseInt(values[i][2]) || 0;
    const amount = totalId * rate;
    paymentSheet.getRange(i + 1, 4).setValue(amount); // Amount column
  }
  
  // Add summary row
  const lastRow = paymentSheet.getLastRow();
  paymentSheet.getRange(lastRow + 2, 1).setValue('SUMMARY');
  paymentSheet.getRange(lastRow + 2, 1).setFontWeight('bold');
  
  const totalIds = values.slice(1).reduce((sum, row) => sum + (parseInt(row[2]) || 0), 0);
  const totalAmount = values.slice(1).reduce((sum, row) => sum + (parseFloat(row[3]) || 0), 0);
  
  paymentSheet.getRange(lastRow + 3, 1).setValue('Total IDs:');
  paymentSheet.getRange(lastRow + 3, 2).setValue(totalIds);
  paymentSheet.getRange(lastRow + 4, 1).setValue('Total Amount:');
  paymentSheet.getRange(lastRow + 4, 2).setValue(totalAmount);
  paymentSheet.getRange(lastRow + 5, 1).setValue('Bad IDs Deducted:');
  paymentSheet.getRange(lastRow + 5, 2).setValue(totalBadIds);
  paymentSheet.getRange(lastRow + 6, 1).setValue('Final Amount:');
  paymentSheet.getRange(lastRow + 6, 2).setValue(totalAmount - (totalBadIds * rate));
  
  // Format summary
  paymentSheet.getRange(lastRow + 2, 1, 5, 2).setFontWeight('bold');
}
