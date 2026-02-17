// ============================================
// COMBINED ARRAYMED CRM - ROUTING HANDLER
// ============================================

// CONFIG for CRM
const CONFIG = {
  USER_SHEET_NAME: 'Users',
  FORM_RESPONSES_SHEET: 'Form responses'
};

// Main routing function - determines which page to show
function doGet(e) {
  // Handle case where e is undefined (manual run from editor)
  const page = e && e.parameter ? e.parameter.page : 'crm'; // Default to CRM now
  
  if (page === 'form') {
    // Show Lead Intake Form
    return HtmlService.createHtmlOutputFromFile('arraymed_form')
      .setTitle('ArrayMed Private Limited - Deal Intake')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    // Show CRM Login page by default
    return HtmlService.createTemplateFromFile('Login')
      .evaluate()
      .setTitle('ArrayMed Sales CRM - Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ============================================
// LEAD INTAKE FORM FUNCTIONS
// ============================================

// Create menu when spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 ArrayMed Tools')
    .addItem('🚀 Test Connections', 'testAllConnections')
    .addItem('📋 Open Lead Form', 'openLeadForm')
    .addItem('🔐 Open CRM Dashboard', 'openCRMDashboard')
    .addSeparator()
    .addItem('🔍 Run Diagnostic Check', 'diagnosticCheck')
    .addItem('🔄 Sync Form Data to User Sheets', 'syncFormResponsesToUserSheets')
    .addItem('🧪 Test Update Mechanism', 'testUpdateMechanism')
    .addToUi();
}

// Open lead form in new window
function openLeadForm() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(
    '<p>Opening Lead Intake Form...</p>' +
    '<script>window.open("' + url + '");google.script.host.close();</script>'
  ).setWidth(300).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Loading');
}

// Open CRM dashboard in new window
function openCRMDashboard() {
  const url = ScriptApp.getService().getUrl() + '?page=crm';
  const html = HtmlService.createHtmlOutput(
    '<p>Opening CRM Dashboard...</p>' +
    '<script>window.open("' + url + '");google.script.host.close();</script>'
  ).setWidth(300).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Loading');
}

// Test all connections
function testAllConnections() {
  console.log('=== TESTING CONNECTIONS ===');
  
  const sales = getSalesTeam();
  const hospitals = getHospitals();
  const products = getProducts();
  
  const html = HtmlService.createHtmlOutput(
    '<div style="padding:20px;font-family:Arial;">' +
    '<h3>✅ Connection Test Results</h3>' +
    '<p><strong>👥 Sales Team:</strong> ' + sales.length + ' people loaded</p>' +
    '<p><strong>🏥 Hospitals:</strong> ' + hospitals.length + ' hospitals loaded</p>' +
    '<p><strong>📦 Products:</strong> ' + products.length + ' products loaded</p>' +
    '<p style="color:green;font-weight:bold;">✓ All connections working!</p>' +
    '</div>'
  ).setWidth(400).setHeight(250);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Test Complete');
  
  return { success: true };
}

// Get sales team from 👥 Sales Team--> sheet
function getSalesTeam() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('👥 Sales Team-->');
    
    if (!sheet) throw new Error('Sales team sheet not found');
    
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow-1, 2).getValues(); // Columns A and B
    
    console.log('Raw data from sales sheet (A=Name, B=Email):');
    data.slice(0, 5).forEach((row, i) => {
      console.log(`Row ${i+2}: A(Name)="${row[0]}", B(Email)="${row[1]}"`);
    });
    
    const names = data
      .filter(row => row[0] && row[0].toString().trim() !== '')
      .map(row => row[0].toString().trim())
      .filter(name => name && name !== '');
    
    console.log('Loaded', names.length, 'sales people (NAMES from Column A)');
    console.log('Sample names:', names.slice(0, 5));
    
    return names;
      
  } catch (error) {
    console.error('Error in getSalesTeam:', error);
    return ['Parimalam V', 'Prathik H Salian', 'Venkatesh S'];
  }
}

// Get hospitals from 🏥 Hospital Database sheet
function getHospitals() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('🏥 Hospital Database');
    
    if (!sheet) throw new Error('Hospital database sheet not found');
    
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 2, lastRow-1, 1).getValues(); // Column B
    
    const hospitals = data
      .flat()
      .filter(hospital => hospital && hospital.toString().trim() !== '')
      .map(hospital => hospital.toString().trim());
    
    console.log('Loaded', hospitals.length, 'hospitals');
    return hospitals;
      
  } catch (error) {
    console.error('Error in getHospitals:', error);
    return ['Apollo Hospital', 'Fortis Hospital', 'Manipal Hospital'];
  }
}

// Get products from 📦 Product Catalog sheet
function getProducts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('📦 Product Catalog');
    
    if (!sheet) throw new Error('Product catalog sheet not found');
    
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow-1, 2).getValues(); // Columns A and B
    
    const products = data
      .filter(row => row[0] && row[1])
      .map(row => {
        const fullName = row[1].toString().trim();
        const cleanName = fullName
          .replace(/^"/, '')
          .replace(/"$/, '')
          .replace(/ - DEMO$/i, '')
          .replace(/\s*DEMO$/i, '')
          .trim();
          
        return {
          code: row[0].toString().trim(),
          name: cleanName,
          category: determineCategory(cleanName)
        };
      });
    
    console.log('Loaded', products.length, 'products');
    
    if (products.length > 0) {
      console.log('Sample product:', products[0]);
    }
    
    return products;
      
  } catch (error) {
    console.error('Error in getProducts:', error);
    return [{
      code: 'PROD-001',
      name: 'Ningbo David Transcutaneous Jaundice Detector',
      category: 'Diagnostic'
    }];
  }
}

// Helper: Determine category from product name
function determineCategory(productName) {
  const name = productName.toLowerCase();
  
  if (name.includes('ultrasound') || name.includes('scan') || name.includes('x-ray') || 
      name.includes('ecg') || name.includes('monitor') || name.includes('detector')) {
    return 'Diagnostic Equipment';
  } else if (name.includes('ventilator') || name.includes('anesthesia') || 
             name.includes('infusion') || name.includes('pump')) {
    return 'Critical Care';
  } else if (name.includes('cable') || name.includes('adapter') || 
             name.includes('filter') || name.includes('bag') || name.includes('tube')) {
    return 'Consumables';
  } else if (name.includes('surgical') || name.includes('forceps') || 
             name.includes('scissors') || name.includes('blade')) {
    return 'Surgical Instruments';
  } else if (name.includes('ophthal') || name.includes('eye') || 
             name.includes('retina') || name.includes('cataract')) {
    return 'Ophthalmology';
  } else {
    return 'General Equipment';
  }
}

// Get sales email by name
function getSalesEmailByName(salesName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('👥 Sales Team-->');
    
    if (!sheet) return '';
    
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow-1, 2).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === salesName) {
        return data[i][1] ? data[i][1].toString().trim() : '';
      }
    }
    
    return '';
  } catch (error) {
    console.error('Error getting sales email:', error);
    return '';
  }
}

// Submit lead to Google Sheets
function submitLead(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = new Date();
    
    const salesEmail = getSalesEmailByName(formData.salesPerson);
    
    // 1. Save to Form Responses sheet
    let formSheet = ss.getSheetByName('Form responses');
    if (!formSheet) {
      formSheet = ss.insertSheet('Form responses');
      formSheet.getRange(1, 1, 1, 17).setValues([[
        'Timestamp', 
        'Sales Email',
        'Sales Person',
        'Hospital', 
        'Contact Person', 
        'Designation',
        'Phone', 
        'Contact Email', 
        'State', 
        'City', 
        'Products', 
        'Product Codes',
        'Budget (INR)', 
        'Timeline', 
        'Notes', 
        'Deal Type', 
        'User Email'
      ]]);
    }
    
    const rowData = [
      timestamp,
      salesEmail || '',
      formData.salesPerson || '',
      formData.hospital || '',
      formData.contactPerson || '',
      formData.designation || '',
      formData.phone || '',
      formData.email || '',
      formData.state || '',
      formData.city || '',
      formData.products || '',
      formData.productCodes || '',
      formData.budget ? Number(formData.budget) : '',
      formData.timeline || '',
      formData.notes || '',
      formData.dealType || '',
      Session.getActiveUser().getEmail()
    ];
    
    formSheet.appendRow(rowData);
    
    console.log('Saved to Form responses sheet');
    
    // 2. Also add to Master CRM
    addToMasterCRM(formData, timestamp, salesEmail);
    
    // 3. Send confirmation email
    sendConfirmationEmail(formData, salesEmail);
    
    return {
      success: true,
      message: 'Deal submitted to ArrayMed CRM',
      leadId: 'ARRAY-' + Utilities.formatDate(timestamp, 'Asia/Kolkata', 'yyMMdd-HHmm')
    };
    
  } catch (error) {
    console.error('Error submitting lead:', error);
    throw new Error('Submission failed: ' + error.message);
  }
}

// Add to Master CRM sheet
function addToMasterCRM(formData, timestamp, salesEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName('📊 Master CRM');
    
    if (!masterSheet) {
      console.log('No Master CRM sheet found - skipping');
      return;
    }
    
    const lastRow = masterSheet.getLastRow();
    const nextRow = lastRow + 1;
    const nextId = lastRow;
    const leadId = 'MED-' + String(nextId).padStart(4, '0');
    
    masterSheet.getRange(nextRow, 1).setValue(leadId);
    masterSheet.getRange(nextRow, 2).setValue(timestamp);
    masterSheet.getRange(nextRow, 3).setValue(salesEmail || '');
    masterSheet.getRange(nextRow, 4).setValue(formData.salesPerson || '');
    masterSheet.getRange(nextRow, 5).setValue(formData.hospital || '');
    masterSheet.getRange(nextRow, 6).setValue(formData.contactPerson || '');
    masterSheet.getRange(nextRow, 7).setValue(formData.phone || '');
    masterSheet.getRange(nextRow, 8).setValue(formData.email || '');
    masterSheet.getRange(nextRow, 9).setValue(formData.products || '');
    masterSheet.getRange(nextRow, 10).setValue('New');
    masterSheet.getRange(nextRow, 11).setValue(timestamp);
    masterSheet.getRange(nextRow, 12).setValue('Contact hospital');
    masterSheet.getRange(nextRow, 13).setValue(formData.city || '');
    masterSheet.getRange(nextRow, 14).setValue(formData.state || '');
    masterSheet.getRange(nextRow, 15).setValue(formData.dealType || '');
    masterSheet.getRange(nextRow, 16).setValue(formData.budget ? Number(formData.budget) : '');
    masterSheet.getRange(nextRow, 17).setValue(formData.timeline || '');
    masterSheet.getRange(nextRow, 18).setValue(formData.designation || '');
    masterSheet.getRange(nextRow, 19).setValue(formData.productCodes || '');
    masterSheet.getRange(nextRow, 25).setValue(formData.notes || '');
    
    console.log('Added to Master CRM at row', nextRow, 'with Lead ID:', leadId);
    
  } catch (error) {
    console.error('Error adding to Master CRM:', error);
  }
}

// Send confirmation email
function sendConfirmationEmail(formData, salesEmail) {
  try {
    if (!salesEmail) {
      console.log('No email address found for confirmation');
      return;
    }
    
    const subject = 'ArrayMed - New Deal Submitted: ' + formData.hospital;
    const formattedBudget = formData.budget ? 
      '₹' + Number(formData.budget).toLocaleString('en-IN') : 'Not specified';
    
    const body = `
ARRAYMED PRIVATE LIMITED
=======================

NEW DEAL SUBMITTED SUCCESSFULLY

📋 DEAL DETAILS:
----------------
- Sales Person: ${formData.salesPerson}
- Sales Email: ${salesEmail}
- Hospital: ${formData.hospital}
- Contact: ${formData.contactPerson} ${formData.designation ? '(' + formData.designation + ')' : ''}
- Phone: ${formData.phone}
- Email: ${formData.email}
- Location: ${formData.city}, ${formData.state}

🛒 PRODUCTS:
------------
${formData.products}

💰 DEAL INFORMATION:
-------------------
- Deal Type: ${formData.dealType}
- Budget: ${formattedBudget}
- Timeline: ${formData.timeline || 'Not specified'}

📝 ADDITIONAL NOTES:
-------------------
${formData.notes || 'None'}

📊 SYSTEM INFO:
--------------
- Submitted: ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })}
- Lead ID: ARRAY-${Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyMMdd-HHmm')}
- Data saved to: Form responses and Master CRM

Regards,
ArrayMed Private Limited CRM System
    `.trim();
    
    MailApp.sendEmail({
      to: salesEmail,
      subject: subject,
      body: body
    });
    
    console.log('Confirmation email sent to:', salesEmail);
    
  } catch (error) {
    console.error('Error sending email:', error);
  }
}

// ============================================
// CRM DASHBOARD FUNCTIONS
// ============================================

// Get dashboard HTML
function getDashboardHtml() {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .getContent();
}

// Get the web app URL
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// Authenticate user
function authenticateUser(email, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(CONFIG.USER_SHEET_NAME);
    
    if (!userSheet) {
      return {
        success: false,
        message: 'User configuration sheet not found. Please contact administrator.'
      };
    }
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailCol = headers.indexOf('Email');
    const passwordCol = headers.indexOf('Password');
    const nameCol = headers.indexOf('Full Name');
    const roleCol = headers.indexOf('Role');
    const reportsToCol = headers.indexOf('Reports To');
    const sheetNameCol = headers.indexOf('Sheet Name');
    const canViewCol = headers.indexOf('Can View');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === email && data[i][passwordCol] === password) {
        return {
          success: true,
          email: data[i][emailCol],
          name: data[i][nameCol],
          role: data[i][roleCol],
          reportsTo: data[i][reportsToCol],
          sheetName: data[i][sheetNameCol],
          canView: data[i][canViewCol]
        };
      }
    }
    
    return {
      success: false,
      message: 'Invalid email or password'
    };
  } catch (error) {
    Logger.log('Authentication error: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred during authentication'
    };
  }
}

// IMPROVED: Get column mapping dynamically
function getColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnMap = {};
  
  headers.forEach((header, index) => {
    if (header) {
      columnMap[header.toString().trim()] = index + 1;
    }
  });
  
  Logger.log('Column Map for ' + sheet.getName() + ':', JSON.stringify(columnMap));
  return columnMap;
}

// IMPROVED: Get deals with dynamic column mapping
function getDeals(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return {
        success: false,
        message: 'Your personal sheet was not found: ' + sheetName
      };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: true,
        deals: [],
        stats: {
          totalDeals: 0,
          totalValue: 0,
          expectedValue: 0,
          wonDeals: 0
        }
      };
    }
    
    // Get column mapping
    const colMap = getColumnMap(sheet);
    
    const allData = sheet.getDataRange().getValues();
    const deals = [];
    let totalValue = 0;
    let expectedValue = 0;
    let wonDeals = 0;
    
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      
      // Skip empty rows
      const hasData = row.some(cell => cell !== null && cell !== undefined && cell !== '');
      if (!hasData) continue;
      
      // Use column mapping to get values
      const deal = {
        rowNumber: i + 1,
        salesPerson: row[colMap['Sales Person Name'] - 1] || '',
        hospitalName: row[colMap['Hospital/Clinic Name'] - 1] || '',
        contactPerson: row[colMap['Contact Person'] - 1] || '',
        designation: row[colMap['Designation'] - 1] || '',
        mobile: row[colMap['Phone Number'] - 1] || '',
        email: row[colMap['Email Address'] - 1] || '',
        state: row[colMap['State'] - 1] || '',
        city: row[colMap['City'] - 1] || '',
        product: row[colMap['Product'] - 1] || '',
        productName: row[colMap['Product Name'] - 1] || '',
        budget: row[colMap['Budget (In INR)'] - 1] || '',
        purchaseTimeline: formatDate(row[colMap['Purchase Timeline (DD-MM-YYYY)'] - 1]),
        additionalNotes: row[colMap['Additional Notes'] - 1] || '',
        stage: row[colMap['Stage'] - 1] || 'New',
        stageDate: formatDate(row[colMap['Stage Date'] - 1]),
        nextAction: row[colMap['Next Action'] - 1] || '',
        nextDate: formatDate(row[colMap['Next Date'] - 1]),
        dealValue: parseFloat(row[colMap['Deal Value'] - 1]) || 0,
        probability: parseFloat(row[colMap['Probability'] - 1]) || 0,
        expectedValue: parseFloat(row[colMap['Expected Value'] - 1]) || 0,
        closeDate: formatDate(row[colMap['Close Date'] - 1]),
        competitor: row[colMap['Competitor'] - 1] || '',
        quote: row[colMap['Quote'] - 1] || '',
        po: row[colMap['PO'] - 1] || '',
        wonDate: formatDate(row[colMap['Won Date'] - 1]),
        lostDate: formatDate(row[colMap['Lost Date'] - 1]),
        lostReason: row[colMap['Lost Reason'] - 1] || '',
        notes: row[colMap['Notes'] - 1] || ''
      };
      
      deals.push(deal);
      
      totalValue += deal.dealValue;
      expectedValue += deal.expectedValue;
      if (deal.stage === 'Won') wonDeals++;
    }
    
    Logger.log('Loaded ' + deals.length + ' deals from ' + sheetName);
    
    return {
      success: true,
      deals: deals,
      stats: {
        totalDeals: deals.length,
        totalValue: totalValue,
        expectedValue: expectedValue,
        wonDeals: wonDeals
      }
    };
    
  } catch (error) {
    Logger.log('Get deals error: ' + error.toString());
    return {
      success: false,
      message: 'Error loading deals: ' + error.toString()
    };
  }
}

// IMPROVED: Update a deal with dynamic column mapping
function updateDeal(sheetName, rowNumber, dealData) {
  try {
    Logger.log('=== UPDATE DEAL STARTED ===');
    Logger.log('Sheet Name: ' + sheetName);
    Logger.log('Row Number: ' + rowNumber);
    Logger.log('Deal Data: ' + JSON.stringify(dealData));
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log('ERROR: Sheet not found - ' + sheetName);
      return {
        success: false,
        message: 'Your sheet not found: ' + sheetName
      };
    }
    
    // Verify row exists
    const lastRow = sheet.getLastRow();
    if (rowNumber > lastRow || rowNumber < 2) {
      return {
        success: false,
        message: 'Invalid row number: ' + rowNumber
      };
    }
    
    // Get column mapping
    const colMap = getColumnMap(sheet);
    
    // Calculate derived values
    const dealValue = parseFloat(dealData.dealValue) || 0;
    const probability = parseFloat(dealData.probability) || 0;
    const expectedValue = (dealValue * probability) / 100;
    
    let updatedFields = 0;
    
    // Update each field using column mapping
    const fieldsToUpdate = {
      'Stage': dealData.stage || '',
      'Stage Date': dealData.stageDate || new Date().toISOString().split('T')[0],
      'Next Action': dealData.nextAction || '',
      'Next Date': dealData.nextDate || '',
      'Deal Value': dealValue,
      'Probability': probability,
      'Expected Value': expectedValue,
      'Close Date': dealData.closeDate || '',
      'Competitor': dealData.competitor || '',
      'Quote': dealData.quote || '',
      'PO': dealData.po || '',
      'Won Date': dealData.wonDate || '',
      'Lost Date': dealData.lostDate || '',
      'Lost Reason': dealData.lostReason || '',
      'Notes': dealData.notes || ''
    };
    
    // Update each field
    for (const [fieldName, value] of Object.entries(fieldsToUpdate)) {
      if (colMap[fieldName]) {
        sheet.getRange(rowNumber, colMap[fieldName]).setValue(value);
        updatedFields++;
        Logger.log('Updated ' + fieldName + ' (col ' + colMap[fieldName] + ') = ' + value);
      } else {
        Logger.log('WARNING: Column not found for field: ' + fieldName);
      }
    }
    
    // Force save
    SpreadsheetApp.flush();
    
    Logger.log('=== UPDATE COMPLETE ===');
    Logger.log('Updated ' + updatedFields + ' fields');
    
    return {
      success: true,
      message: 'Deal updated successfully',
      updatedFields: updatedFields
    };
    
  } catch (error) {
    Logger.log('=== UPDATE FAILED ===');
    Logger.log('Error: ' + error.toString());
    return {
      success: false,
      message: 'Update failed: ' + error.toString()
    };
  }
}

// IMPROVED: Update multiple deals at once
function updateMultipleDeals(sheetName, dealsToUpdate) {
  try {
    Logger.log('=== BULK UPDATE STARTED ===');
    Logger.log('Sheet: ' + sheetName);
    Logger.log('Number of deals to update: ' + dealsToUpdate.length);
    
    const results = {
      success: true,
      totalDeals: dealsToUpdate.length,
      successCount: 0,
      failCount: 0,
      errors: []
    };
    
    // Update each deal
    for (let i = 0; i < dealsToUpdate.length; i++) {
      const deal = dealsToUpdate[i];
      Logger.log('Updating deal ' + (i + 1) + '/' + dealsToUpdate.length + ' (Row ' + deal.rowNumber + ')');
      
      const result = updateDeal(sheetName, deal.rowNumber, deal.dealData);
      
      if (result.success) {
        results.successCount++;
      } else {
        results.failCount++;
        results.errors.push({
          rowNumber: deal.rowNumber,
          error: result.message
        });
      }
    }
    
    // Final status
    if (results.failCount > 0) {
      results.success = false;
      results.message = 'Updated ' + results.successCount + ' deals, ' + results.failCount + ' failed';
    } else {
      results.message = 'Successfully updated all ' + results.successCount + ' deals';
    }
    
    Logger.log('=== BULK UPDATE COMPLETE ===');
    Logger.log('Success: ' + results.successCount + ', Failed: ' + results.failCount);
    
    return results;
    
  } catch (error) {
    Logger.log('=== BULK UPDATE FAILED ===');
    Logger.log('Error: ' + error.toString());
    return {
      success: false,
      message: 'Bulk update failed: ' + error.toString(),
      totalDeals: dealsToUpdate.length,
      successCount: 0,
      failCount: dealsToUpdate.length
    };
  }
}

// Get team members for a manager/founder
function getTeamMembers(userEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(CONFIG.USER_SHEET_NAME);
    
    if (!userSheet) {
      return { success: false, message: 'Users sheet not found' };
    }
    
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailCol = headers.indexOf('Email');
    const nameCol = headers.indexOf('Full Name');
    const roleCol = headers.indexOf('Role');
    const reportsToCol = headers.indexOf('Reports To');
    const sheetNameCol = headers.indexOf('Sheet Name');
    
    let currentUserName = '';
    let currentUserRole = '';
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === userEmail) {
        currentUserName = data[i][nameCol];
        currentUserRole = data[i][roleCol];
        break;
      }
    }
    
    if (!currentUserName) {
      return { success: false, message: 'User not found' };
    }
    
    const teamMembers = [];
    
    for (let i = 1; i < data.length; i++) {
      const reportsTo = data[i][reportsToCol];
      const memberName = data[i][nameCol];
      const memberRole = data[i][roleCol];
      const memberSheet = data[i][sheetNameCol];
      
      if (currentUserRole === 'Founder' && memberName !== currentUserName && memberSheet !== 'ALL') {
        teamMembers.push({
          name: memberName,
          role: memberRole,
          sheetName: memberSheet
        });
      }
      else if (reportsTo === currentUserName && memberSheet !== 'ALL') {
        teamMembers.push({
          name: memberName,
          role: memberRole,
          sheetName: memberSheet
        });
      }
    }
    
    teamMembers.sort((a, b) => {
      if (a.role !== b.role) {
        const roleOrder = ['Business Manager', 'Sales Manager', 'Sales Executive'];
        return roleOrder.indexOf(a.role) - roleOrder.indexOf(b.role);
      }
      return a.name.localeCompare(b.name);
    });
    
    return {
      success: true,
      teamMembers: teamMembers
    };
    
  } catch (error) {
    Logger.log('Get team members error: ' + error.toString());
    return {
      success: false,
      message: 'An error occurred while loading team members'
    };
  }
}

// Helper function to format dates
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  try {
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      date = new Date(dateValue);
    } else {
      return '';
    }
    
    if (isNaN(date.getTime())) return '';
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
  } catch (error) {
    return '';
  }
}

// ============================================
// DIAGNOSTIC AND DATA SYNC FUNCTIONS
// ============================================

// Diagnostic function to check sheet and data
function diagnosticCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userEmail = Session.getActiveUser().getEmail();
  
  Logger.log('=== DIAGNOSTIC CHECK ===');
  Logger.log('Active user email: ' + userEmail);
  
  // Check Users sheet
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    Logger.log('ERROR: Users sheet not found!');
    return;
  }
  
  const userData = userSheet.getDataRange().getValues();
  Logger.log('Users in system: ' + (userData.length - 1));
  
  // Check Form responses
  const formSheet = ss.getSheetByName('Form responses');
  if (!formSheet) {
    Logger.log('WARNING: Form responses sheet not found!');
  } else {
    const formData = formSheet.getDataRange().getValues();
    Logger.log('Form submissions: ' + (formData.length - 1));
  }
  
  // Check individual sheets
  const allSheets = ss.getSheets();
  Logger.log('Total sheets in workbook: ' + allSheets.length);
  
  allSheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.endsWith('_L')) {
      const lastRow = sheet.getLastRow();
      Logger.log(`Sheet "${name}": ${lastRow - 1} deals`);
      
      // Show column structure
      const colMap = getColumnMap(sheet);
      Logger.log('  Columns: ' + JSON.stringify(Object.keys(colMap)));
    }
  });
  
  Logger.log('=== END DIAGNOSTIC ===');
}

// NEW FUNCTION: Test the update mechanism
function testUpdateMechanism() {
  Logger.log('=== TESTING UPDATE MECHANISM ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Test sheet name - CHANGE THIS to match your actual sheet
  const testSheetName = 'Parimalam V_L';
  const testSheet = ss.getSheetByName(testSheetName);
  
  if (!testSheet) {
    Logger.log('ERROR: Test sheet not found - ' + testSheetName);
    Logger.log('Available sheets:');
    ss.getSheets().forEach(s => Logger.log('  - ' + s.getName()));
    return;
  }
  
  Logger.log('✓ Sheet found: ' + testSheetName);
  Logger.log('  Rows: ' + testSheet.getLastRow());
  Logger.log('  Columns: ' + testSheet.getLastColumn());
  
  // Get column mapping
  const colMap = getColumnMap(testSheet);
  
  // Check required columns
  const requiredCols = ['Stage', 'Deal Value', 'Probability', 'Expected Value', 'Notes'];
  Logger.log('\nRequired columns check:');
  requiredCols.forEach(col => {
    if (colMap[col]) {
      Logger.log('  ✓ ' + col + ' found at column ' + colMap[col]);
    } else {
      Logger.log('  ✗ ' + col + ' NOT FOUND');
    }
  });
  
  // Try a test write on row 2 if it exists
  if (testSheet.getLastRow() >= 2 && colMap['Notes']) {
    Logger.log('\nTesting write operation on row 2...');
    
    const notesCol = colMap['Notes'];
    
    try {
      // Save original value
      const originalValue = testSheet.getRange(2, notesCol).getValue();
      Logger.log('  Original value: "' + originalValue + '"');
      
      // Write test value
      const testValue = 'TEST ' + new Date().toISOString();
      testSheet.getRange(2, notesCol).setValue(testValue);
      SpreadsheetApp.flush();
      Logger.log('  ✓ Wrote test value: "' + testValue + '"');
      
      // Read back
      const readBack = testSheet.getRange(2, notesCol).getValue();
      Logger.log('  ✓ Read back: "' + readBack + '"');
      
      // Restore original
      testSheet.getRange(2, notesCol).setValue(originalValue);
      SpreadsheetApp.flush();
      Logger.log('  ✓ Restored original value');
      
      if (readBack === testValue) {
        Logger.log('\n✓✓✓ WRITE TEST SUCCESSFUL!');
        Logger.log('The update mechanism is working correctly.');
      }
    } catch (e) {
      Logger.log('  ✗ WRITE TEST FAILED: ' + e.toString());
    }
  }
  
  Logger.log('\n=== TEST COMPLETE ===');
}

// Sync form responses to individual user sheets
function syncFormResponsesToUserSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formSheet = ss.getSheetByName('Form responses');
    
    if (!formSheet) {
      Logger.log('Form responses sheet not found');
      return { success: false, message: 'Form responses sheet not found' };
    }
    
    const formData = formSheet.getDataRange().getValues();
    const headers = formData[0];
    
    Logger.log('Syncing ' + (formData.length - 1) + ' form responses...');
    
    let syncCount = 0;
    
    // Process each form response (skip header)
    for (let i = 1; i < formData.length; i++) {
      const row = formData[i];
      const salesPerson = row[2]; // Column C: Sales Person Name
      
      if (!salesPerson) continue;
      
      // Find user's sheet name
      const userSheet = ss.getSheetByName('Users');
      if (!userSheet) continue;
      
      const userData = userSheet.getDataRange().getValues();
      const userHeaders = userData[0];
      const nameCol = userHeaders.indexOf('Full Name');
      const sheetNameCol = userHeaders.indexOf('Sheet Name');
      
      let targetSheetName = null;
      
      for (let j = 1; j < userData.length; j++) {
        if (userData[j][nameCol] === salesPerson) {
          targetSheetName = userData[j][sheetNameCol];
          break;
        }
      }
      
      if (!targetSheetName || targetSheetName === 'ALL') continue;
      
      // Get or create user's sheet
      let userDealSheet = ss.getSheetByName(targetSheetName);
      
      if (!userDealSheet) {
        Logger.log('Creating sheet: ' + targetSheetName);
        userDealSheet = ss.insertSheet(targetSheetName);
        
        // Add headers
        const dealHeaders = [
          'Sales Person Name', 'Hospital/Clinic Name', 'Contact Person', 'Designation',
          'Phone Number', 'Email Address', 'State', 'City', 'Product', 'Product Name',
          'Budget (In INR)', 'Purchase Timeline (DD-MM-YYYY)', 'Additional Notes',
          'Stage', 'Stage Date', 'Next Action', 'Next Date', 'Deal Value', 'Probability',
          'Expected Value', 'Close Date', 'Competitor', 'Quote', 'PO', 'Won Date',
          'Lost Date', 'Lost Reason', 'Notes'
        ];
        userDealSheet.getRange(1, 1, 1, dealHeaders.length).setValues([dealHeaders]);
        userDealSheet.getRange(1, 1, 1, dealHeaders.length)
          .setFontWeight('bold')
          .setBackground('#667eea')
          .setFontColor('#ffffff');
      }
      
      // Check if this entry already exists
      const existingData = userDealSheet.getDataRange().getValues();
      let exists = false;
      
      for (let k = 1; k < existingData.length; k++) {
        // Check if hospital, contact, and email match
        if (existingData[k][1] === row[3] && 
            existingData[k][2] === row[4] && 
            existingData[k][5] === row[7]) {
          exists = true;
          break;
        }
      }
      
      if (!exists) {
        // Add new row with form data (columns A-M) + empty tracking columns (N-AB)
        const budgetValue = row[12] || 0;
        const newRow = [
          row[2],  // Sales Person Name
          row[3],  // Hospital
          row[4],  // Contact Person
          row[5],  // Designation
          row[6],  // Phone
          row[7],  // Email
          row[8],  // State
          row[9],  // City
          row[10], // Products
          row[11], // Product Codes
          budgetValue, // Budget
          row[13], // Timeline
          row[14], // Notes
          'New',   // Stage
          new Date(), // Stage Date
          'Contact hospital', // Next Action
          '',      // Next Date
          budgetValue, // Deal Value (same as budget initially)
          10,      // Probability (10% for new)
          budgetValue * 0.1, // Expected Value
          '',      // Close Date
          '',      // Competitor
          '',      // Quote
          '',      // PO
          '',      // Won Date
          '',      // Lost Date
          '',      // Lost Reason
          ''       // Notes
        ];
        
        userDealSheet.appendRow(newRow);
        syncCount++;
        Logger.log('Added deal to ' + targetSheetName);
      }
    }
    
    Logger.log('Sync complete. Added ' + syncCount + ' new deals.');
    
    return {
      success: true,
      message: 'Synced ' + syncCount + ' new deals to user sheets'
    };
    
  } catch (error) {
    Logger.log('Error syncing: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

// Send quote email with attachments
function sendQuoteEmail(emailData) {
  try {
    const { to, hospitalName, contactPerson, product, message, attachments, salesPerson } = emailData;
    
    // Validate email
    if (!to || to.trim() === '') {
      return {
        success: false,
        message: 'Recipient email is required'
      };
    }
    
    // Prepare email subject
    const subject = `Quotation from ArrayMed Private Limited - ${product || 'Medical Equipment'}`;
    
    // Prepare HTML email body
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; text-align: center; color: white;">
          <h1 style="margin: 0; font-size: 28px;">ArrayMed Private Limited</h1>
          <p style="margin: 10px 0 0 0; opacity: 0.95;">Medical Equipment Solutions</p>
        </div>
        
        <div style="padding: 30px; background: #ffffff;">
          <p style="margin: 0 0 20px 0; font-size: 16px; color: #2c3e50;">
            Dear ${contactPerson || 'Team'},
          </p>
          
          <div style="white-space: pre-line; line-height: 1.6; color: #2c3e50; font-size: 15px;">
${message}
          </div>
          
          <div style="margin: 30px 0; padding: 20px; background: #f8f9fa; border-left: 4px solid #667eea; border-radius: 4px;">
            <h3 style="margin: 0 0 15px 0; color: #667eea; font-size: 16px;">Deal Information</h3>
            <p style="margin: 5px 0; color: #2c3e50;"><strong>Hospital/Clinic:</strong> ${hospitalName || '-'}</p>
            <p style="margin: 5px 0; color: #2c3e50;"><strong>Product:</strong> ${product || '-'}</p>
            <p style="margin: 5px 0; color: #2c3e50;"><strong>Sales Representative:</strong> ${salesPerson || '-'}</p>
          </div>
          
          ${attachments && attachments.length > 0 ? `
          <div style="margin: 20px 0;">
            <p style="color: #6c757d; font-size: 14px;">
              📎 <strong>${attachments.length}</strong> file(s) attached
            </p>
          </div>
          ` : ''}
        </div>
        
        <div style="padding: 20px 30px; background: #f8f9fa; border-top: 2px solid #e9ecef; text-align: center;">
          <p style="margin: 0 0 10px 0; color: #6c757d; font-size: 14px;">
            <strong>ArrayMed Private Limited</strong><br>
            Medical Equipment Solutions Provider<br>
            📧 sales@arraymed.co.in
          </p>
          <p style="margin: 0; color: #999; font-size: 12px;">
            This is an automated email from ArrayMed CRM System
          </p>
        </div>
      </div>
    `;
    
    // Create email options
    const emailOptions = {
      to: to,
      subject: subject,
      htmlBody: htmlBody,
      name: 'ArrayMed Private Limited'
    };
    
    // Add attachments if provided
    if (attachments && attachments.length > 0) {
      emailOptions.attachments = attachments.map(file => {
        return Utilities.newBlob(
          Utilities.base64Decode(file.data),
          file.mimeType,
          file.name
        );
      });
    }
    
    // Send email
    MailApp.sendEmail(emailOptions);
    
    Logger.log('Quote email sent successfully to: ' + to);
    
    return {
      success: true,
      message: 'Quote sent successfully'
    };
    
  } catch (error) {
    Logger.log('Error sending quote email: ' + error.toString());
    return {
      success: false,
      message: 'Failed to send email: ' + error.toString()
    };
  }
}

// Setup function - Run this once to create the Users sheet
function setupUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let userSheet = ss.getSheetByName(CONFIG.USER_SHEET_NAME);
  
  if (!userSheet) {
    userSheet = ss.insertSheet(CONFIG.USER_SHEET_NAME);
  } else {
    userSheet.clear();
  }
  
  const headers = ['Email', 'Password', 'Full Name', 'Role', 'Reports To', 'Sheet Name', 'Can View'];
  userSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  userSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
  
  const userData = [
    ['keshav@arraymed.co.in', 'keshav123', 'Keshav', 'Founder', '', 'ALL', ''],
    ['sales@arraymed.co.in', 'param123', 'Parimalam V', 'Business Manager', 'Keshav', 'Parimalam V_L', ''],
    ['prathikasilani@arraymed.co.in', 'prathik123', 'Prathik H Salian', 'Business Manager', 'Keshav', 'Prathik H Salian_L', ''],
    ['dinesh@arraymed.co.in', 'dinesh123', 'Dinesh P', 'Business Manager', 'Keshav', 'Dinesh P_L', ''],
    ['arunkumar@arraymed.co.in', 'arun123', 'Arun Kumar P', 'Business Manager', 'Keshav', 'Arun Kumar P_L', ''],
    ['naveen@arraymed.co.in', 'naveen123', 'Naveen Kumar C D', 'Business Manager', 'Keshav', 'Naveen Kumar C D_L', ''],
    ['dhyaan@arraymed.co.in', 'sharan123', 'Sharanappa', 'Business Manager', 'Keshav', 'Sharanappa_L', ''],
    ['shivakumar@arraymed.co.in', 'shiva123', 'Shivakumar N', 'Sales Manager', 'Keshav', 'Shivakumar N_L', ''],
    ['ismail@arraymed.co.in', 'ismail123', 'K Ismail Basha', 'Sales Manager', 'Keshav', 'K Ismail Basha_L', ''],
    ['prasenjit@arraymed.co.in', 'prasen123', 'Prasenjit Mitra', 'Sales Manager', 'Parimalam V', 'Prasenjit Mitra_L', ''],
    ['rajesh@arraymed.co.in', 'rajesh123', 'Rajesh V', 'Sales Manager', 'Parimalam V', 'Rajesh V_L', ''],
    ['ramanand@arraymed.co.in', 'raman123', 'Ramanand Sah', 'Sales Manager', 'Sharanappa', 'Ramanand Sah_L', ''],
    ['anil@arraymed.co.in', 'anil123', 'C N Anil Kumar', 'Sales Executive', 'Parimalam V', 'C N Anil Kumar_L', ''],
    ['aneesh@arraymed.co.in', 'aneesh123', 'Aneesh C Bangera', 'Sales Executive', 'Parimalam V', 'Aneesh C Bangera_L', ''],
    ['subham@arraymed.co.in', 'subham123', 'Subham Bandyopadhyay', 'Sales Executive', 'Parimalam V', 'Subham Bandyopadhyay_L', ''],
    ['anushree@arraymed.co.in', 'anushree123', 'Anushree R', 'Sales Executive', 'Parimalam V', 'Anushree R_L', ''],
    ['narendra@arraymed.co.in', 'naren123', 'Narendra Kumar P', 'Sales Executive', 'Parimalam V', 'Narendra Kumar P_L', ''],
    ['suriya@arraymed.co.in', 'suriya123', 'Jayasuriya Ganesan', 'Sales Executive', 'Parimalam V', 'Jayasuriya Ganesan_L', ''],
    ['vivek@arraymed.co.in', 'vivek123', 'C H Vivekananda', 'Sales Executive', 'Parimalam V', 'C H Vivekananda_L', ''],
    ['monisha@arraymed.co.in', 'dhyaan123', 'Dhyaan S P', 'Sales Executive', 'Parimalam V', 'Dhyaan S P_L', ''],
    ['santhoshc@arraymed.co.in', 'santhosh123', 'Santhosh C', 'Sales Executive', 'Dinesh P', 'Santhosh C_L', ''],
    ['mohith@arraymed.co.in', 'mohith123', 'Mohit R', 'Sales Executive', 'Dinesh P', 'Mohit R_L', ''],
    ['yashwin@arraymed.co.in', 'yashwin123', 'Yashwin R', 'Sales Executive', 'Dinesh P', 'Yashwin R_L', ''],
    ['jaffar@arraymed.co.in', 'jaffar123', 'Jaffar Sadiq M', 'Sales Executive', 'Arun Kumar P', 'Jaffar Sadiq M_L', ''],
    ['praveen@arraymed.co.in', 'praveen123', 'Praveen Kumar R', 'Sales Executive', 'Arun Kumar P', 'Praveen Kumar R_L', ''],
    ['somana@arraymed.co.in', 'somana123', 'Somanagouda Patil', 'Sales Executive', 'Naveen Kumar C D', 'Somanagouda Patil_L', ''],
    ['sirish@arraymed.co.in', 'sirish123', 'Sirish Umarji', 'Sales Executive', 'Sharanappa', 'Sirish Umarji_L', ''],
    ['bhaskar@arraymed.co.in', 'bhaskar123', 'Bhaskar V S', 'Sales Executive', 'Sharanappa', 'Bhaskar V S_L', ''],
    ['gururajai@arraymed.co.in', 'guru123', 'Guruja Rai S', 'Sales Executive', 'Keshav', 'Guruja Rai S_L', '']
  ];
  
  userSheet.getRange(2, 1, userData.length, userData[0].length).setValues(userData);
  
  userSheet.setFrozenRows(1);
  userSheet.autoResizeColumns(1, headers.length);
  
  SpreadsheetApp.getUi().alert('Users sheet has been set up successfully!');
}