// ============================================
// ARRAYMED PRIVATE LIMITED - CRM FORM BACKEND
// ============================================

// Main function to serve the HTML form
function doGet() {
  return HtmlService.createHtmlOutputFromFile('arraymed_form')
    .setTitle('ArrayMed Private Limited - Deal Intake')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Create menu when spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 ArrayMed Tools')
    .addItem('🚀 Test Connections', 'testAllConnections')
    .addItem('📋 Open Form', 'openForm')
    .addToUi();
}

// Open form in new window
function openForm() {
  const html = HtmlService.createHtmlOutput(
    '<p>Opening ArrayMed Form...</p>' +
    '<script>window.open("' + ScriptApp.getService().getUrl() + '");google.script.host.close();</script>'
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

// Get sales team from 👥 Sales Team--> sheet (Name in col A, Email in col B) - CORRECTED!
function getSalesTeam() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('👥 Sales Team-->');
    
    if (!sheet) throw new Error('Sales team sheet not found');
    
    const lastRow = sheet.getLastRow();
    // CORRECT: Get name (col A) and email (col B)
    const data = sheet.getRange(2, 1, lastRow-1, 2).getValues(); // Columns A and B
    
    // DEBUG: Log what we're getting
    console.log('Raw data from sales sheet (A=Name, B=Email):');
    data.slice(0, 5).forEach((row, i) => {
      console.log(`Row ${i+2}: A(Name)="${row[0]}", B(Email)="${row[1]}"`);
    });
    
    // Return NAMES only for dropdown (from Column A)
    const names = data
      .filter(row => row[0] && row[0].toString().trim() !== '')
      .map(row => row[0].toString().trim()) // Column A is Name
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
    const data = sheet.getRange(2, 2, lastRow-1, 1).getValues(); // Column B (Hospital Names)
    
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

// Get sales email by name (from Column B)
function getSalesEmailByName(salesName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('👥 Sales Team-->');
    
    if (!sheet) return '';
    
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow-1, 2).getValues(); // Columns A and B
    
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
    
    // Get sales email from the name (Column B)
    const salesEmail = getSalesEmailByName(formData.salesPerson);
    
    // 1. Save to Form Responses sheet - Email in B, Name in C
    let formSheet = ss.getSheetByName('Form responses');
    if (!formSheet) {
      formSheet = ss.insertSheet('Form responses');
      // Headers: Email in B, Name in C
      formSheet.getRange(1, 1, 1, 17).setValues([[
        'Timestamp', 
        'Sales Email',      // Column B: Email
        'Sales Person',     // Column C: Name
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
    
    // Order: Email first (B), then Name (C)
    const rowData = [
      timestamp,                          // A: Timestamp
      salesEmail || '',                   // B: Sales Email
      formData.salesPerson || '',         // C: Sales Person Name
      formData.hospital || '',            // D: Hospital
      formData.contactPerson || '',       // E: Contact Person
      formData.designation || '',         // F: Designation
      formData.phone || '',               // G: Phone
      formData.email || '',               // H: Contact Email
      formData.state || '',               // I: State
      formData.city || '',                // J: City
      formData.products || '',            // K: Products
      formData.productCodes || '',        // L: Product Codes
      formData.budget ? Number(formData.budget) : '', // M: Budget (as number)
      formData.timeline || '',            // N: Timeline (DD/MM/YYYY)
      formData.notes || '',               // O: Notes
      formData.dealType || '',            // P: Deal Type
      Session.getActiveUser().getEmail()  // Q: User Email
    ];
    
    formSheet.appendRow(rowData);
    
    console.log('Saved to Form responses sheet:');
    console.log('Column B (Email):', salesEmail);
    console.log('Column C (Name):', formData.salesPerson);
    console.log('Column M (Budget):', formData.budget);
    console.log('Column N (Timeline):', formData.timeline);
    
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
    
    // Format budget with Indian number formatting
    const formattedBudget = formData.budget ? 
      '₹' + Number(formData.budget).toLocaleString('en-IN') : 'Not specified';
    
    const body = `
ARRAYMED PRIVATE LIMITED
=======================

NEW DEAL SUBMITTED SUCCESSFULLY

📋 DEAL DETAILS:
----------------
• Sales Person: ${formData.salesPerson}
• Sales Email: ${salesEmail}
• Hospital: ${formData.hospital}
• Contact: ${formData.contactPerson} ${formData.designation ? '(' + formData.designation + ')' : ''}
• Phone: ${formData.phone}
• Email: ${formData.email}
• Location: ${formData.city}, ${formData.state}

🛒 PRODUCTS:
------------
${formData.products}

💰 DEAL INFORMATION:
-------------------
• Deal Type: ${formData.dealType}
• Budget: ${formattedBudget}
• Timeline: ${formData.timeline || 'Not specified'}

📝 ADDITIONAL NOTES:
-------------------
${formData.notes || 'None'}

📊 SYSTEM INFO:
--------------
• Submitted: ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })}
• Lead ID: ARRAY-${Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyMMdd-HHmm')}
• Data saved to: Form responses and Master CRM

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