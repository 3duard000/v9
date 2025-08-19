/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Sheet Manager - SheetManager.gs
 * 
 * This module handles all sheet creation, formatting, column widths,
 * dropdowns, and comprehensive conditional formatting.
 * 
 * FIXED: Added error handling and validation for null values
 */

const SheetManager = {

  /**
   * Create all required sheets with headers and formatting
   */
  createRequiredSheets() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Validate that constants are available
      if (!SHEET_NAMES || !HEADERS) {
        throw new Error('SHEET_NAMES or HEADERS constants are not defined. Check Config.gs file.');
      }
      
      console.log('Creating required sheets with validation...');
      
      // Create all main sheets (including Guest Rooms)
      const mainSheets = ['TENANT', 'BUDGET', 'MAINTENANCE', 'GUEST_ROOMS'];
      
      mainSheets.forEach(key => {
        try {
          console.log(`\nProcessing sheet: ${key}`);
          
          // Validate sheet name
          const sheetName = SHEET_NAMES[key];
          if (!sheetName) {
            throw new Error(`Sheet name not found for key: ${key}. Check SHEET_NAMES in Config.gs`);
          }
          console.log(`Sheet name: "${sheetName}"`);
          
          // Validate headers
          const headers = HEADERS[key];
          if (!headers || !Array.isArray(headers) || headers.length === 0) {
            throw new Error(`Headers not found or invalid for key: ${key}. Check HEADERS in Config.gs`);
          }
          console.log(`Headers found: ${headers.length} columns`);
          
          // Check for null/undefined values in headers
          const nullIndex = headers.findIndex(header => header === null || header === undefined || header === '');
          if (nullIndex !== -1) {
            throw new Error(`Null/undefined header found at index ${nullIndex} for sheet ${key}`);
          }
          console.log(`All headers valid for ${key}`);
          
          let sheet = spreadsheet.getSheetByName(sheetName);
          
          if (!sheet) {
            console.log(`Creating new sheet: ${sheetName}`);
            sheet = spreadsheet.insertSheet(sheetName);
          }
          
          // Set headers if the sheet is empty or has different headers
          const lastRow = sheet.getLastRow();
          if (lastRow === 0) {
            // Sheet is completely empty, add headers
            console.log(`Adding headers to ${sheetName}...`);
            this._addHeaders(sheet, headers);
            this._formatSheet(sheet, key);
            console.log(`✅ Sheet "${sheetName}" created/updated with headers and formatting`);
          } else {
            // Check if headers match
            this._updateHeadersIfNeeded(sheet, headers, key);
          }
          
        } catch (sheetError) {
          console.error(`Error processing sheet ${key}:`, sheetError);
          throw new Error(`Failed to create/update sheet ${key}: ${sheetError.message}`);
        }
      });
      
      console.log('✅ All required sheets created successfully');
      
    } catch (error) {
      console.error('Error in createRequiredSheets:', error);
      throw new Error(`Sheet creation failed: ${error.message}`);
    }
  },

  /**
   * Add headers to a sheet with validation
   * @private
   */
  _addHeaders(sheet, headers) {
    try {
      // Validate inputs
      if (!sheet) {
        throw new Error('Sheet is null or undefined');
      }
      if (!headers || !Array.isArray(headers) || headers.length === 0) {
        throw new Error('Headers array is null, undefined, or empty');
      }
      
      // Check for null values in headers array
      headers.forEach((header, index) => {
        if (header === null || header === undefined) {
          throw new Error(`Header at index ${index} is null or undefined`);
        }
      });
      
      console.log(`Setting headers: [${headers.join(', ')}]`);
      console.log(`Range: A1:${String.fromCharCode(65 + headers.length - 1)}1`);
      
      // Set the headers
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row with stronger blue
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1c4587'); // Stronger blue
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      console.log(`✅ Headers added successfully to ${sheet.getName()}`);
      
    } catch (error) {
      console.error('Error in _addHeaders:', error);
      throw new Error(`Failed to add headers: ${error.message}`);
    }
  },

  /**
   * Update headers if they don't match expected headers
   * @private
   */
  _updateHeadersIfNeeded(sheet, headers, key) {
    try {
      const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      const existingHeadersString = existingHeaders.join('');
      const expectedHeadersString = headers.join('');
      
      if (existingHeadersString !== expectedHeadersString) {
        console.log(`Updating headers for ${sheet.getName()}`);
        this._addHeaders(sheet, headers);
        this._formatSheet(sheet, key);
        console.log(`✅ Sheet "${sheet.getName()}" headers updated`);
      } else {
        console.log(`Headers already correct for ${sheet.getName()}`);
      }
    } catch (error) {
      console.error('Error updating headers:', error);
      throw new Error(`Failed to update headers: ${error.message}`);
    }
  },

  /**
   * Apply formatting to a sheet (column widths, freezing, dropdowns, conditional formatting)
   * @private
   */
  _formatSheet(sheet, sheetKey) {
    try {
      console.log(`Formatting sheet: ${sheetKey}`);
      
      // Set column widths
      this._setColumnWidths(sheet, sheetKey);
      
      // Freeze header row
      sheet.setFrozenRows(1);
      
      // Add dropdowns and conditional formatting
      this._addDropdownsAndFormatting(sheet, sheetKey);
      
      console.log(`✅ Formatting completed for ${sheetKey}`);
      
    } catch (error) {
      console.error(`Error formatting sheet ${sheetKey}:`, error);
      throw new Error(`Failed to format sheet: ${error.message}`);
    }
  },

  /**
   * Set column widths for a sheet
   * @private
   */
  _setColumnWidths(sheet, sheetKey) {
    try {
      const columnWidths = COLUMN_WIDTHS[sheetKey];
      if (!columnWidths) {
        console.log(`No column widths defined for ${sheetKey}, skipping...`);
        return;
      }
      
      // Apply column widths
      for (let i = 0; i < columnWidths.length; i++) {
        sheet.setColumnWidth(i + 1, columnWidths[i]);
      }
      
      console.log(`Column widths set for ${sheet.getName()}`);
      
    } catch (error) {
      console.error(`Error setting column widths for ${sheetKey}:`, error);
      // Don't throw here - column widths are not critical
    }
  },

  /**
   * Add dropdown menus and conditional formatting to a sheet
   * @private
   */
  _addDropdownsAndFormatting(sheet, sheetKey) {
    try {
      switch(sheetKey) {
        case 'TENANT':
          this._formatTenantSheet(sheet);
          break;
        case 'GUEST_ROOMS':
          this._formatGuestRoomsSheet(sheet);
          break;
        case 'BUDGET':
          this._formatBudgetSheet(sheet);
          break;
        case 'MAINTENANCE':
          this._formatMaintenanceSheet(sheet);
          break;
        default:
          console.log(`No specific formatting defined for ${sheetKey}`);
      }
    } catch (error) {
      console.error(`Error adding dropdowns/formatting for ${sheetKey}:`, error);
      // Don't throw here - formatting is not critical for basic functionality
    }
  },

  /**
   * Format Tenant sheet with dropdowns and conditional formatting
   * @private
   */
  _formatTenantSheet(sheet) {
    try {
      // Room Status dropdown (Column I - index 9)
      const roomStatusRange = sheet.getRange('I2:I1000');
      this._addDataValidation(roomStatusRange, DROPDOWN_OPTIONS.TENANT.ROOM_STATUS);
      
      // Payment Status dropdown (Column K - index 11)
      const paymentStatusRange = sheet.getRange('K2:K1000');
      this._addDataValidation(paymentStatusRange, DROPDOWN_OPTIONS.TENANT.PAYMENT_STATUS);
      
      // Conditional formatting for Room Status (Column I)
      const roomStatusRules = [
        this._createConditionalRule('Occupied', COLORS.LIGHT_GREEN, roomStatusRange),
        this._createConditionalRule('Vacant', COLORS.LIGHT_YELLOW, roomStatusRange),
        this._createConditionalRule('Maintenance', COLORS.LIGHT_RED, roomStatusRange),
        this._createConditionalRule('Reserved', COLORS.LIGHT_BLUE, roomStatusRange)
      ];
      
      // Conditional formatting for Payment Status (Column K)
      const paymentStatusRules = [
        this._createConditionalRule('Current', COLORS.LIGHT_GREEN, paymentStatusRange),
        this._createConditionalRule('Late', COLORS.LIGHT_ORANGE, paymentStatusRange),
        this._createConditionalRule('Overdue', COLORS.LIGHT_RED, paymentStatusRange),
        this._createConditionalRule('Partial', COLORS.LIGHT_YELLOW, paymentStatusRange)
      ];
      
      // Combine all rules
      const allTenantRules = roomStatusRules.concat(paymentStatusRules);
      sheet.setConditionalFormatRules(allTenantRules);
      
      console.log('✅ Tenant sheet dropdowns and formatting added');
      
    } catch (error) {
      console.error('Error formatting tenant sheet:', error);
    }
  },

  /**
   * Format Guest Rooms sheet with dropdowns and conditional formatting
   * @private
   */
  _formatGuestRoomsSheet(sheet) {
    try {
      // Room Type dropdown (Column D)
      this._addDataValidation(sheet.getRange('D2:D1000'), DROPDOWN_OPTIONS.GUEST_ROOMS.ROOM_TYPE);
      
      // Payment Status dropdown (Column O - index 15)
      const paymentStatusRange = sheet.getRange('O2:O1000');
      this._addDataValidation(paymentStatusRange, DROPDOWN_OPTIONS.GUEST_ROOMS.PAYMENT_STATUS);
      
      // Booking Status dropdown (Column P - index 16)
      const bookingStatusRange = sheet.getRange('P2:P1000');
      this._addDataValidation(bookingStatusRange, DROPDOWN_OPTIONS.GUEST_ROOMS.BOOKING_STATUS);
      
      // Source dropdown (Column Q - index 17)
      this._addDataValidation(sheet.getRange('Q2:Q1000'), DROPDOWN_OPTIONS.GUEST_ROOMS.SOURCE);
      
      // Conditional formatting for Payment Status (Column O)
      const paymentStatusRules = [
        this._createConditionalRule('Paid', COLORS.LIGHT_GREEN, paymentStatusRange),
        this._createConditionalRule('Pending', COLORS.LIGHT_YELLOW, paymentStatusRange),
        this._createConditionalRule('Deposit Paid', COLORS.LIGHT_BLUE, paymentStatusRange),
        this._createConditionalRule('Cancelled', COLORS.LIGHT_RED, paymentStatusRange),
        this._createConditionalRule('Refunded', COLORS.LIGHT_ORANGE, paymentStatusRange)
      ];
      
      // Conditional formatting for Booking Status (Column P)
      const bookingStatusRules = [
        this._createConditionalRule('Confirmed', COLORS.LIGHT_GREEN, bookingStatusRange),
        this._createConditionalRule('Checked-In', COLORS.LIGHT_BLUE, bookingStatusRange),
        this._createConditionalRule('Checked-Out', '#f3f3f3', bookingStatusRange),
        this._createConditionalRule('Reserved', COLORS.LIGHT_YELLOW, bookingStatusRange),
        this._createConditionalRule('Cancelled', COLORS.LIGHT_RED, bookingStatusRange),
        this._createConditionalRule('No-Show', COLORS.LIGHT_RED, bookingStatusRange)
      ];
      
      // Combine all rules
      const allGuestRules = paymentStatusRules.concat(bookingStatusRules);
      sheet.setConditionalFormatRules(allGuestRules);
      
      console.log('✅ Guest Rooms sheet dropdowns and formatting added');
      
    } catch (error) {
      console.error('Error formatting guest rooms sheet:', error);
    }
  },

  /**
   * Format Budget sheet with dropdowns and conditional formatting
   * @private
   */
  _formatBudgetSheet(sheet) {
    try {
      // Type dropdown (Column B)
      const typeRange = sheet.getRange('B2:B1000');
      this._addDataValidation(typeRange, DROPDOWN_OPTIONS.BUDGET.TYPE);
      
      // Category dropdown (Column E)
      this._addDataValidation(sheet.getRange('E2:E1000'), DROPDOWN_OPTIONS.BUDGET.CATEGORY);
      
      // Payment Method dropdown (Column F)
      this._addDataValidation(sheet.getRange('F2:F1000'), DROPDOWN_OPTIONS.BUDGET.PAYMENT_METHOD);
      
      // Amount range (Column D) - for conditional formatting
      const amountRange = sheet.getRange('D2:D1000');
      
      // Custom conditional formatting based on Type column
      // Income - format Amount column with light green background
      const incomeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B2="Income"')
        .setBackground(COLORS.LIGHT_GREEN)
        .setFontColor('#000000')
        .setRanges([amountRange])
        .build();
      
      // Expense - format Amount column with light red background  
      const expenseRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B2="Expense"')
        .setBackground(COLORS.LIGHT_RED)
        .setFontColor('#000000')
        .setRanges([amountRange])
        .build();
      
      // Conditional formatting for Type column itself
      const typeRules = [
        this._createConditionalRule('Income', COLORS.LIGHT_GREEN, typeRange),
        this._createConditionalRule('Expense', COLORS.LIGHT_RED, typeRange)
      ];
      
      // Combine all rules
      const allBudgetRules = [incomeRule, expenseRule].concat(typeRules);
      sheet.setConditionalFormatRules(allBudgetRules);
      
      console.log('✅ Budget sheet dropdowns and formatting added');
      
    } catch (error) {
      console.error('Error formatting budget sheet:', error);
    }
  },

  /**
   * Format Maintenance sheet with dropdowns and conditional formatting
   * @private
   */
  _formatMaintenanceSheet(sheet) {
    try {
      // Issue Type dropdown (Column D)
      this._addDataValidation(sheet.getRange('D2:D1000'), DROPDOWN_OPTIONS.MAINTENANCE.ISSUE_TYPE);
      
      // Priority dropdown (Column E)
      const priorityRange = sheet.getRange('E2:E1000');
      this._addDataValidation(priorityRange, DROPDOWN_OPTIONS.MAINTENANCE.PRIORITY);
      
      // Status dropdown (Column J)
      const statusRange = sheet.getRange('J2:J1000');
      this._addDataValidation(statusRange, DROPDOWN_OPTIONS.MAINTENANCE.STATUS);
      
      // Conditional formatting for Priority (Column E)
      const priorityRules = [
        this._createConditionalRule('Low', COLORS.LIGHT_GREEN, priorityRange),
        this._createConditionalRule('Medium', COLORS.LIGHT_YELLOW, priorityRange),
        this._createConditionalRule('High', COLORS.LIGHT_ORANGE, priorityRange),
        this._createConditionalRule('Emergency', COLORS.LIGHT_RED, priorityRange)
      ];
      
      // Conditional formatting for Status (Column J)
      const statusRules = [
        this._createConditionalRule('Completed', COLORS.LIGHT_GREEN, statusRange),
        this._createConditionalRule('In Progress', COLORS.LIGHT_BLUE, statusRange),
        this._createConditionalRule('Pending', COLORS.LIGHT_YELLOW, statusRange),
        this._createConditionalRule('Cancelled', '#f3f3f3', statusRange),
        this._createConditionalRule('On Hold', COLORS.LIGHT_ORANGE, statusRange)
      ];
      
      // Combine all maintenance rules
      const allMaintenanceRules = priorityRules.concat(statusRules);
      sheet.setConditionalFormatRules(allMaintenanceRules);
      
      console.log('✅ Maintenance sheet dropdowns and formatting added');
      
    } catch (error) {
      console.error('Error formatting maintenance sheet:', error);
    }
  },

  /**
   * Add data validation (dropdown) to a range with error handling
   * @private
   */
  _addDataValidation(range, options) {
    try {
      if (!options || !Array.isArray(options) || options.length === 0) {
        console.log('No dropdown options provided, skipping validation');
        return;
      }
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(options)
        .setAllowInvalid(false)
        .build();
      range.setDataValidation(rule);
      
    } catch (error) {
      console.error('Error adding data validation:', error);
      // Don't throw - dropdowns are not critical for basic functionality
    }
  },

  /**
   * Create a conditional formatting rule with black text for better readability
   * @private
   */
  _createConditionalRule(textValue, backgroundColor, range) {
    try {
      return SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(textValue)
        .setBackground(backgroundColor)
        .setFontColor('#000000') // Black text for better readability
        .setRanges([range])
        .build();
    } catch (error) {
      console.error('Error creating conditional rule:', error);
      return null;
    }
  }
};
