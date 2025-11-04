// ============================================
// LogOperations.gs - Work with History sheet
// ============================================

// Log user action
function logAction(invoiceId, userInfo, action, oldStatus, newStatus, details) {
  try {
    const sheet = getOrCreateHistorySheet();
    
    const invoice = invoiceId ? getInvoiceById(invoiceId) : null;
    const invoiceNumber = invoice ? invoice.number : '';
    
    const timestamp = new Date().toLocaleString('ru-RU', {timeZone: 'Asia/Almaty'});
    
    const rowData = [
      timestamp,
      userInfo.name || '',
      userInfo.role || '',
      action,
      invoiceNumber,
      oldStatus || '',
      newStatus || '',
      details || ''
    ];
    
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    Logger.log('📝 Action logged: ' + action);
    
  } catch (error) {
    Logger.log('❌ Error logging action: ' + error);
  }
}

// Get logs (only for admin)
function getLogs(filters) {
  try {
    const sheet = getOrCreateHistorySheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return [];
    }
    
    const logs = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const log = {
        timestamp: String(row[0] || ''),
        user: String(row[1] || ''),
        role: String(row[2] || ''),
        action: String(row[3] || ''),
        invoiceNumber: String(row[4] || ''),
        oldStatus: String(row[5] || ''),
        newStatus: String(row[6] || ''),
        details: String(row[7] || '')
      };
      
      // Apply filters if provided
      if (filters) {
        if (filters.user && log.user !== filters.user) continue;
        if (filters.action && log.action !== filters.action) continue;
        if (filters.invoiceNumber && log.invoiceNumber !== filters.invoiceNumber) continue;
      }
      
      logs.push(log);
    }
    
    Logger.log('📋 Loaded ' + logs.length + ' logs');
    return logs;
    
  } catch (error) {
    Logger.log('❌ Error loading logs: ' + error);
    return [];
  }
}

// Helper: Get or create History sheet
function getOrCreateHistorySheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(CONFIG.HISTORY_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('📋 Creating History sheet');
      sheet = spreadsheet.insertSheet(CONFIG.HISTORY_SHEET_NAME);
      createHistoryHeaders(sheet);
    }
    
    return sheet;
  } catch (error) {
    Logger.log('❌ Error accessing History sheet: ' + error);
    throw error;
  }
}

// Helper: Create History headers
function createHistoryHeaders(sheet) {
  const headers = [
    'timestamp',
    'user',
    'role',
    'action',
    'invoiceNumber',
    'oldStatus',
    'newStatus',
    'details'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ea4335');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
}
