// ============================================
// SupplierOperations.gs - Work with Suppliers sheet
// ============================================

// Get all suppliers (for autocomplete)
function getSuppliers() {
  try {
    const sheet = getOrCreateSuppliersSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return [];
    }
    
    const suppliers = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      suppliers.push({
        name: String(row[0] || ''),
        bin: String(row[1] || ''),
        count: parseInt(row[2]) || 0,
        lastUsed: String(row[3] || '')
      });
    }
    
    // Sort by usage count (most used first)
    suppliers.sort((a, b) => b.count - a.count);
    
    Logger.log('📋 Loaded ' + suppliers.length + ' suppliers');
    return suppliers;
    
  } catch (error) {
    Logger.log('❌ Error loading suppliers: ' + error);
    return [];
  }
}

// Update or create supplier template
function updateSupplierTemplate(name, bin) {
  try {
    if (!name) return;
    
    const sheet = getOrCreateSuppliersSheet();
    const data = sheet.getDataRange().getValues();
    
    let found = false;
    let rowIndex = -1;
    
    // Search for existing supplier
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === name) {
        found = true;
        rowIndex = i + 1;
        break;
      }
    }
    
    const now = new Date().toLocaleString('ru-RU', {timeZone: 'Asia/Almaty'});
    
    if (found) {
      // Update existing
      const currentCount = parseInt(data[rowIndex-1][2]) || 0;
      sheet.getRange(rowIndex, 3).setValue(currentCount + 1); // Increase count
      sheet.getRange(rowIndex, 4).setValue(now); // Update lastUsed
      
      // Update BIN if provided and different
      if (bin && bin !== data[rowIndex-1][1]) {
        sheet.getRange(rowIndex, 2).setValue(bin);
      }
      
      Logger.log('✅ Supplier updated: ' + name);
      
    } else {
      // Create new
      const rowData = [
        name,
        bin || '',
        1,
        now
      ];
      
      const nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
      
      Logger.log('✅ Supplier created: ' + name);
    }
    
  } catch (error) {
    Logger.log('❌ Error updating supplier: ' + error);
  }
}

// Helper: Get or create Suppliers sheet
function getOrCreateSuppliersSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(CONFIG.SUPPLIERS_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('📋 Creating Suppliers sheet');
      sheet = spreadsheet.insertSheet(CONFIG.SUPPLIERS_SHEET_NAME);
      createSuppliersHeaders(sheet);
    }
    
    return sheet;
  } catch (error) {
    Logger.log('❌ Error accessing Suppliers sheet: ' + error);
    throw error;
  }
}

// Helper: Create Suppliers headers
function createSuppliersHeaders(sheet) {
  const headers = [
    'name',
    'bin',
    'count',
    'lastUsed'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
}
