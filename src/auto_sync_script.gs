const CONFIG = {
  FORM_ID: 'Change with your ID of G.Form',
  SHEET_ID: 'Change with your ID of G.Sheet',
};

  const DROPDOWNS = [
    {
      sheetName: 'category',
      column: 'A',
      startRow: 2,
      questionTitle: 'Transaction Name'
    },
    {
      sheetName: 'category',
      column: 'B',
      startRow: 2,
      questionTitle: 'Category'
    }
  ];
  
  // ==============
  // MENU CREATION
  // ==============
  
  function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🔄 Form Sync')
      .addItem('📥 Sync Dropdowns Now', 'syncFromMenu')
      .addItem('📊 Show Status', 'showStatusDialog')
      .addSeparator()
      .addItem('⚙️ Initial Setup', 'initialSetup')
      .addItem('🔍 Check Configuration', 'validateConfig')
      .addToUi();
      
    console.log('Menu created. Look for "🔄 Form Sync" in your menu bar.');
  }
  
  // ==============
  // INITIAL SETUP
  // ==============
  
  function initialSetup() {
    const ui = SpreadsheetApp.getUi();
    
    console.log('=== INITIAL SETUP ===\n');
    
    let configValid = true;
    let setupMessage = '';
    
    try {
      const form = FormApp.openById(CONFIG.FORM_ID);
      setupMessage += `✅ Form found: "${form.getTitle()}"\n`;
      console.log(`✅ Form found: "${form.getTitle()}"`);
    } catch(e) {
      setupMessage += '❌ ERROR: Cannot access form. Check your FORM_ID\n';
      console.log('❌ ERROR: Cannot access form. Check your FORM_ID');
      configValid = false;
    }
    
    try {
      const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      setupMessage += `✅ Spreadsheet found: "${spreadsheet.getName()}"\n`;
      console.log(`✅ Spreadsheet found: "${spreadsheet.getName()}"`);
    } catch(e) {
      setupMessage += '❌ ERROR: Cannot access spreadsheet. Check your SHEET_ID\n';
      console.log('❌ ERROR: Cannot access spreadsheet. Check your SHEET_ID');
      configValid = false;
    }
    
    if (!configValid) {
      ui.alert('Setup Failed', setupMessage + '\nPlease fix the configuration errors.', ui.ButtonSet.OK);
      return;
    }
    
    createStatusDisplay();
    
    setupMessage += '\n✅ Setup complete!\n\nUse the menu "🔄 Form Sync → Sync Dropdowns Now" to sync your dropdowns.';
    
    ui.alert('Setup Complete', setupMessage, ui.ButtonSet.OK);
    console.log('Setup complete. Use the menu to sync dropdowns.');
  }
  
  /**
   * Creates status display in Form Sync Status sheet
   */
  function createStatusDisplay() {
    try {
      const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      let sheet = spreadsheet.getSheetByName('Form Sync Status');
      let sheetCreated = false;
      
      // Create Form Sync Status sheet if it doesn't exist
      if (!sheet) {
        sheet = spreadsheet.insertSheet('Form Sync Status');
        console.log('Created Form Sync Status sheet');
        sheetCreated = true;
        SpreadsheetApp.flush();
        // Add small delay to ensure sheet is fully created
        Utilities.sleep(500);
      }
      
      // Clear existing content first
      sheet.clear();
      
      // Set column widths BEFORE adding content
      sheet.setColumnWidth(1, 200);
      sheet.setColumnWidth(2, 350);
      
      // === TITLE SECTION (Row 1) ===
      const titleCell = sheet.getRange('A1');
      titleCell.setValue('📋 FORM SYNC STATUS');
      
      // Merge AFTER setting the value
      sheet.getRange('A1:B1').merge();
      
      // Apply formatting to merged cell
      const mergedTitle = sheet.getRange('A1:B1');
      mergedTitle
        .setFontSize(14)
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      
      // Set row height for title
      sheet.setRowHeight(1, 40);
      
      // === STATUS SECTION (Rows 3-5) ===
      // Status row
      sheet.getRange('A3').setValue('Status:').setFontWeight('bold');
      sheet.getRange('B3')
        .setValue('⚪ Ready to sync')
        .setBackground('#f5f5f5');
      
      // Last Sync row
      sheet.getRange('A4').setValue('Last Sync:').setFontWeight('bold');
      sheet.getRange('B4')
        .setValue('Never')
        .setBackground('#f5f5f5');
      
      // Dropdowns row
      sheet.getRange('A5').setValue('Dropdowns:').setFontWeight('bold');
      sheet.getRange('B5')
        .setValue(`${DROPDOWNS.length} configured`)
        .setBackground('#f5f5f5');
      
      // Add border to status section
      sheet.getRange('A3:B5').setBorder(
        true, true, true, true, true, true,
        '#cccccc', SpreadsheetApp.BorderStyle.SOLID
      );
      
      // === INSTRUCTIONS SECTION (Rows 7-10) ===
      sheet.getRange('A7')
        .setValue('HOW TO SYNC:')
        .setFontWeight('bold')
        .setFontSize(11);
      
      // Merge instructions title
      sheet.getRange('A7:B7').merge();
      sheet.getRange('A7:B7')
        .setBackground('#e8f5e9')
        .setHorizontalAlignment('left');
      
      // Individual instruction steps
      const instructions = [
        '1. Click "🔄 Form Sync" in the menu bar above',
        '2. Select "📥 Sync Dropdowns Now"',
        '3. Confirm to start syncing'
      ];
      
      instructions.forEach((instruction, index) => {
        const row = 8 + index;
        sheet.getRange(`A${row}:B${row}`).merge();
        sheet.getRange(`A${row}:B${row}`)
          .setValue(instruction)
          .setWrap(true)
          .setVerticalAlignment('middle');
        sheet.setRowHeight(row, 30);
      });
      
      // Add subtle border to instructions section
      sheet.getRange('A7:B10').setBorder(
        true, true, true, true, false, false,
        '#cccccc', SpreadsheetApp.BorderStyle.SOLID
      );
      
      // === GENERAL FORMATTING ===
      // Set default row heights
      sheet.setRowHeight(2, 10); // Empty row after title
      sheet.setRowHeight(6, 10); // Empty row before instructions
      
      // Set text alignment for label column
      sheet.getRange('A3:A5').setHorizontalAlignment('right');
      
      // Final flush to ensure all changes are saved
      SpreadsheetApp.flush();
      
      console.log('✅ Status display created successfully in Form Sync Status sheet');
      return sheetCreated;
      
    } catch(error) {
      console.error('❌ Error creating status display:', error.toString());
      throw error;
    }
  }
  
  // ====================
  // MAIN SYNC FUNCTIONS
  // ====================
  
  function syncFromMenu() {
    const ui = SpreadsheetApp.getUi();
    
    const result = ui.alert(
      '🔄 Sync Dropdowns',
      `This will update ${DROPDOWNS.length} form dropdown(s) from your sheets.\n\nContinue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (result == ui.Button.YES) {
      updateStatus('🔄 Syncing...', null, null);
      const syncResult = syncDropdowns();
      
      if (syncResult.success) {
        ui.alert(
          '✅ Sync Complete', 
          `Successfully updated ${syncResult.count} dropdown(s) with ${syncResult.totalOptions} total options.\n\nCheck your form to see the updates.`,
          ui.ButtonSet.OK
        );
      } else {
        ui.alert(
          '⚠️ Sync Completed with Issues',
          `Updated ${syncResult.count} dropdown(s).\n\nIssues found:\n${syncResult.errors.join('\n')}\n\nCheck the Apps Script logs for details.`,
          ui.ButtonSet.OK
        );
      }
    }
  }
  
  function syncDropdowns() {
    console.log('=== STARTING SYNC ===');
    
    const form = FormApp.openById(CONFIG.FORM_ID);
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    
    let successCount = 0;
    let totalOptions = 0;
    let errors = [];
    
    DROPDOWNS.forEach(config => {
      try {
        const sheet = spreadsheet.getSheetByName(config.sheetName);
        
        if (!sheet) {
          const error = `Sheet "${config.sheetName}" not found`;
          console.log(`❌ ERROR: ${error}`);
          errors.push(error);
          return;
        }
        
        const lastRow = sheet.getLastRow();
        if (lastRow < config.startRow) {
          console.log(`⚠️ No data in sheet "${config.sheetName}" starting from row ${config.startRow}`);
          return;
        }
        
        const col = config.column.charCodeAt(0) - 64;
        const values = sheet.getRange(config.startRow, col, lastRow - config.startRow + 1, 1)
          .getValues()
          .flat()
          .filter(v => v !== '' && v !== null);
        
        if (values.length === 0) {
          console.log(`⚠️ No values found in ${config.sheetName} column ${config.column}`);
          return;
        }
        
        const uniqueValues = [...new Set(values)];
        const items = form.getItems();
        let found = false;
        
        items.forEach(item => {
          if (item.getTitle() === config.questionTitle && item.getType() === FormApp.ItemType.LIST) {
            item.asListItem().setChoiceValues(uniqueValues);
            successCount++;
            totalOptions += uniqueValues.length;
            found = true;
            console.log(`✅ Updated "${config.questionTitle}" with ${uniqueValues.length} options`);
          }
        });
        
        if (!found) {
          const error = `Dropdown "${config.questionTitle}" not found in form`;
          console.log(`⚠️ WARNING: ${error}`);
          errors.push(error);
        }
        
      } catch(e) {
        const error = `Error updating ${config.questionTitle}: ${e.toString()}`;
        console.log(`❌ ERROR: ${error}`);
        errors.push(error);
      }
    });
    
    const timestamp = new Date().toLocaleString();
    if (successCount > 0) {
      updateStatus('✅ Synced', timestamp, `${successCount} dropdowns, ${totalOptions} options`);
    } else {
      updateStatus('❌ Sync failed', timestamp, 'Check logs for errors');
    }
    
    console.log(`\n=== SYNC COMPLETE ===`);
    console.log(`✅ Success: ${successCount}/${DROPDOWNS.length} dropdowns updated`);
    console.log(`📊 Total: ${totalOptions} options`);
    
    return {
      success: errors.length === 0,
      count: successCount,
      totalOptions: totalOptions,
      errors: errors
    };
  }
  
  function updateStatus(status, timestamp, details) {
    try {
      const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      const sheet = spreadsheet.getSheetByName('Form Sync Status');
      
      if (!sheet) {
        console.log('⚠️ Form Sync Status sheet not found, creating it...');
        createStatusDisplay();
        return;
      }
      
      const statusCell = sheet.getRange('B3');
      statusCell.setValue(status);
      
      if (status.includes('✅')) {
        statusCell.setBackground('#e8f5e9').setFontColor('#2e7d32');
      } else if (status.includes('🔄')) {
        statusCell.setBackground('#fff3e0').setFontColor('#f57c00');
      } else if (status.includes('❌')) {
        statusCell.setBackground('#ffebee').setFontColor('#c62828');
      } else {
        statusCell.setBackground('#f5f5f5').setFontColor('#000000');
      }
      
      if (timestamp) {
        sheet.getRange('B4')
          .setValue(timestamp)
          .setBackground('#e3f2fd')
          .setFontColor('#1565c0');
      }
      
      if (details) {
        sheet.getRange('B5')
          .setValue(details)
          .setBackground('#f5f5f5');
      }
      
      SpreadsheetApp.flush();
      
    } catch(e) {
      console.error('❌ Could not update status display:', e.toString());
    }
  }
  
  // ============================================
  // HELPER FUNCTIONS
  // ============================================
  
  function showStatusDialog() {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = spreadsheet.getSheetByName('Form Sync Status');
    
    let statusMessage = '📊 CURRENT STATUS\n\n';
    
    if (sheet) {
      const status = sheet.getRange('B3').getValue();
      const lastSync = sheet.getRange('B4').getValue();
      const details = sheet.getRange('B5').getValue();
      
      statusMessage += `Status: ${status}\n`;
      statusMessage += `Last Sync: ${lastSync}\n`;
      statusMessage += `Details: ${details}\n`;
    } else {
      statusMessage += 'No status information available.\nRun Initial Setup first.';
    }
    
    statusMessage += '\n\nCONFIGURED DROPDOWNS:\n';
    DROPDOWNS.forEach((config, i) => {
      statusMessage += `${i+1}. "${config.questionTitle}" ← ${config.sheetName} (Col ${config.column})\n`;
    });
    
    ui.alert('Form Sync Status', statusMessage, ui.ButtonSet.OK);
  }
  
  function validateConfig() {
    const ui = SpreadsheetApp.getUi();
    let message = '🔍 CONFIGURATION CHECK\n\n';
    let hasErrors = false;
    
    try {
      const form = FormApp.openById(CONFIG.FORM_ID);
      message += `✅ Form: "${form.getTitle()}"\n`;
      
      const items = form.getItems();
      const dropdowns = items.filter(item => item.getType() === FormApp.ItemType.LIST);
      message += `   Found ${dropdowns.length} dropdown(s) in form\n\n`;
      
    } catch(e) {
      message += '❌ Form: Cannot access (check FORM_ID)\n\n';
      hasErrors = true;
    }
    
    try {
      const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      message += `✅ Spreadsheet: "${spreadsheet.getName()}"\n`;
      
      message += '\nSheet Check:\n';
      DROPDOWNS.forEach(config => {
        const sheet = spreadsheet.getSheetByName(config.sheetName);
        if (sheet) {
          const lastRow = sheet.getLastRow();
          message += `✅ "${config.sheetName}" - ${lastRow} rows\n`;
        } else {
          message += `❌ "${config.sheetName}" - NOT FOUND\n`;
          hasErrors = true;
        }
      });
      
    } catch(e) {
      message += '❌ Spreadsheet: Cannot access (check SHEET_ID)\n';
      hasErrors = true;
    }
    
    message += '\n' + (hasErrors ? '⚠️ Please fix the errors above.' : '✅ All configurations valid!');
    
    ui.alert('Configuration Check', message, ui.ButtonSet.OK);
  }
  
  function testSync() {
    console.log('=== MANUAL TEST ===');
    const result = syncDropdowns();
    console.log('Test complete. Result:', result);
  }
