/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Manual Tenant Application Panel - ManualTenantApplication.gs
 * 
 * This module provides a user-friendly panel interface to manually enter tenant applications
 * and directly assign them to rooms without needing the Google Form process.
 */

const ManualTenantApplication = {

  /**
   * Show the manual tenant application entry panel
   */
  showManualApplicationPanel() {
    try {
      console.log('Opening Manual Tenant Application Panel...');
      
      const availableRooms = this._getAvailableRooms();
      
      if (availableRooms.length === 0) {
        SpreadsheetApp.getUi().alert(
          'No Available Rooms',
          'No available rooms found for new tenant assignment. All rooms appear to be occupied, in maintenance, or reserved.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }
      
      const html = this._generateManualApplicationHTML(availableRooms);
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(900)
        .setHeight(750)
        .setTitle('📝 Manual Tenant Application');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Manual Tenant Application');
      
    } catch (error) {
      console.error('Error showing manual application panel:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load manual application panel: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Process the manual tenant application and add to tenant sheet
   */
  processManualApplication(applicationData) {
    try {
      console.log('Processing manual application:', applicationData);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
      
      if (!tenantSheet) {
        throw new Error('Tenant sheet not found');
      }
      
      // Parse the application data
      const data = JSON.parse(applicationData);
      
      // Find the row with the selected room number
      const roomRowIndex = this._findRoomRow(tenantSheet, data.roomNumber);
      if (roomRowIndex === -1) {
        throw new Error(`Room ${data.roomNumber} not found in tenant sheet`);
      }
      
      // Check if room is actually available
      const currentRoomStatus = this._getRoomStatus(tenantSheet, roomRowIndex);
      if (currentRoomStatus.toLowerCase() === 'occupied') {
        throw new Error(`Room ${data.roomNumber} is currently occupied and cannot be assigned to a new tenant`);
      }
      
      // Update the existing row with tenant information - matching exact tenant sheet columns
      const tenantRow = [
        data.roomNumber,                     // Room Number
        data.rentalPrice,                    // Rental Price
        data.negotiatedPrice || '',          // Negotiated Price
        data.tenantName,                     // Current Tenant Name
        data.tenantEmail,                    // Tenant Email
        data.tenantPhone,                    // Tenant Phone
        data.moveInDate,                     // Move-In Date
        data.securityDeposit,                // Security Deposit Paid
        'Occupied',                          // Room Status
        data.lastPaymentDate || '',          // Last Payment Date
        data.paymentStatus || 'Current',     // Payment Status
        data.moveOutPlanned || '',           // Move-Out Date (Planned)
        data.emergencyContact || '',         // Emergency Contact
        data.leaseEndDate || '',             // Lease End Date
        `Manual entry on ${new Date().toLocaleDateString()}. ${data.notes || ''}`.trim() // Notes
      ];
      
      // Update the existing row
      tenantSheet.getRange(roomRowIndex, 1, 1, tenantRow.length).setValues([tenantRow]);
      
      console.log(`Manual application processed successfully for ${data.tenantName} in Room ${data.roomNumber}`);
      return `✅ Tenant added successfully! ${data.tenantName} has been assigned to Room ${data.roomNumber} and the room is now marked as occupied.`;
      
    } catch (error) {
      console.error('Error processing manual application:', error);
      throw new Error('Failed to process manual application: ' + error.message);
    }
  },

  /**
   * Get available rooms from tenant sheet
   * @private
   */
  _getAvailableRooms() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const tenantSheet = spreadsheet.getSheetByName(SHEET_NAMES.TENANT);
      
      if (!tenantSheet || tenantSheet.getLastRow() <= 1) {
        return [];
      }
      
      const data = tenantSheet.getDataRange().getValues();
      const headers = data[0];
      const rooms = [];
      
      const roomNumberCol = headers.indexOf('Room Number');
      const rentalPriceCol = headers.indexOf('Rental Price');
      const statusCol = headers.indexOf('Room Status');
      const tenantNameCol = headers.indexOf('Current Tenant Name');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[roomNumberCol]) {
          const roomNumber = row[roomNumberCol].toString();
          const rentalPrice = row[rentalPriceCol] || '$0';
          const status = (row[statusCol] || 'Available').toString().toLowerCase();
          const tenantName = row[tenantNameCol] || '';
          
          // Only include available rooms (vacant or available, not occupied/maintenance/reserved)
          if (status === 'vacant' || status === 'available' || (status !== 'occupied' && !tenantName)) {
            let statusDisplay = '';
            switch (status) {
              case 'vacant':
              case 'available':
                statusDisplay = '(Available)';
                break;
              case 'maintenance':
                statusDisplay = '(Maintenance - Available Soon)';
                break;
              case 'reserved':
                statusDisplay = '(Reserved)';
                break;
              default:
                statusDisplay = '(Available)';
                break;
            }
            
            rooms.push({
              number: roomNumber,
              price: rentalPrice,
              status: status,
              display: `Room ${roomNumber} - ${rentalPrice}/month ${statusDisplay}`,
              available: true
            });
          }
        }
      }
      
      // Sort by room number
      rooms.sort((a, b) => a.number.localeCompare(b.number));
      
      console.log(`Found ${rooms.length} available rooms for manual assignment`);
      return rooms;
      
    } catch (error) {
      console.error('Error getting available rooms:', error);
      return [];
    }
  },

  /**
   * Find the row index for a specific room number
   * @private
   */
  _findRoomRow(sheet, roomNumber) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const roomNumberCol = headers.indexOf('Room Number');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][roomNumberCol].toString() === roomNumber.toString()) {
        return i + 1; // Return 1-based row index
      }
    }
    
    return -1; // Room not found
  },

  /**
   * Get current room status
   * @private
   */
  _getRoomStatus(sheet, rowIndex) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol = headers.indexOf('Room Status');
    
    if (statusCol !== -1) {
      return sheet.getRange(rowIndex, statusCol + 1).getValue() || 'Available';
    }
    
    return 'Available';
  },

  /**
   * Generate HTML for the manual application panel
   * @private
   */
  _generateManualApplicationHTML(availableRooms) {
    const roomsJson = JSON.stringify(availableRooms);
    
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #6f42c1; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .form-container {
            background: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 25px;
            margin-bottom: 20px;
        }
        .form-section {
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 1px solid #e0e0e0;
        }
        .form-section:last-child {
            border-bottom: none;
            margin-bottom: 0;
        }
        .section-title {
            font-size: 16px;
            font-weight: bold;
            color: #6f42c1;
            margin-bottom: 20px;
            padding-bottom: 8px;
            border-bottom: 2px solid #6f42c1;
        }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; color: #333; }
        .form-group input, .form-group select, .form-group textarea { 
            width: 100%; 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            margin-bottom: 8px;
            box-sizing: border-box;
        }
        .form-group textarea {
            resize: vertical;
            min-height: 80px;
        }
        .form-group small {
            color: #666; 
            font-size: 11px; 
            margin-top: 8px; 
            display: block;
            line-height: 1.4;
        }
        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 25px;
            margin-bottom: 25px;
        }
        .form-row-three {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 20px;
            margin-bottom: 25px;
        }
        .btn { 
            background: #6f42c1; 
            color: white; 
            padding: 12px 25px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 14px;
            font-weight: bold;
        }
        .btn:hover { background: #5a32a3; }
        .btn-secondary { 
            background: #6c757d; 
            margin-left: 10px;
        }
        .btn-secondary:hover { background: #5a6268; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .required { color: #dc3545; }
        .no-rooms { text-align: center; color: #666; margin: 50px 0; }
        .currency-input { position: relative; }
        .currency-symbol { position: absolute; left: 15px; top: 50%; transform: translateY(-50%); color: #666; z-index: 1; }
        .currency-input input { padding-left: 25px; }
        .room-info {
            background: #e7e3ff;
            border: 1px solid #c7b3ff;
            border-radius: 5px;
            padding: 15px;
            margin: 15px 0;
            display: none;
        }
        .room-info h4 { margin: 0 0 10px 0; color: #6f42c1; }
    </style>
</head>
<body>
    <div class="header">
        <h2>📝 Manual Tenant Application</h2>
        <p>Enter tenant information matching the tenant sheet columns</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    ${availableRooms.length === 0 ? `
        <div class="no-rooms">
            <h3>No Available Rooms</h3>
            <p>All rooms are currently occupied, in maintenance, or reserved. Please free up a room first before adding a new tenant.</p>
        </div>
    ` : `
        <div class="form-container">
            <form id="manualApplicationForm">
                <div class="form-section">
                    <div class="section-title">Tenant Information</div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>Current Tenant Name <span class="required">*</span>:</label>
                            <input type="text" id="tenantName" required>
                            <small>Full legal name of the tenant</small>
                        </div>
                        <div class="form-group">
                            <label>Tenant Email <span class="required">*</span>:</label>
                            <input type="text" id="tenantEmail" required>
                            <small>Enter email address (validation will be relaxed)</small>
                        </div>
                    </div>
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label>Tenant Phone <span class="required">*</span>:</label>
                            <input type="text" id="tenantPhone" required>
                            <small>Phone number with area code</small>
                        </div>
                        <div class="form-group">
                            <label>Emergency Contact:</label>
                            <input type="text" id="emergencyContact" placeholder="Name and phone number">
                        </div>
                    </div>
                </div>
                
                <div class="form-section">
                    <div class="section-title">Room Assignment & Dates</div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>Room Number <span class="required">*</span>:</label>
                            <select id="roomNumber" onchange="updateRoomInfo()" required>
                                <option value="">-- Select Room --</option>
                                ${availableRooms.map(room => `
                                    <option value="${room.number}" data-price="${room.price}">${room.display}</option>
                                `).join('')}
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Move-In Date <span class="required">*</span>:</label>
                            <input type="date" id="moveInDate" required>
                        </div>
                    </div>
                    
                    <div id="roomInfo" class="room-info">
                        <h4>Room Information</h4>
                        <p><strong>Rental Price:</strong> <span id="rentalPriceDisplay">$0</span>/month</p>
                        <p><strong>Status:</strong> Available for assignment</p>
                    </div>
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label>Lease End Date:</label>
                            <input type="date" id="leaseEndDate">
                            <small>When the lease expires</small>
                        </div>
                        <div class="form-group">
                            <label>Move-Out Date (Planned):</label>
                            <input type="date" id="moveOutPlanned">
                            <small>If tenant has mentioned a planned move-out date</small>
                        </div>
                    </div>
                </div>
                
                <div class="form-section">
                    <div class="section-title">Financial Information</div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>Negotiated Price:</label>
                            <div class="currency-input">
                                <span class="currency-symbol">$</span>
                                <input type="text" id="negotiatedPrice" placeholder="1200">
                            </div>
                            <small>Leave blank to use the room's rental price</small>
                        </div>
                        <div class="form-group">
                            <label>Security Deposit Paid <span class="required">*</span>:</label>
                            <div class="currency-input">
                                <span class="currency-symbol">$</span>
                                <input type="text" id="securityDeposit" required placeholder="1200">
                            </div>
                        </div>
                    </div>
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label>Last Payment Date:</label>
                            <input type="date" id="lastPaymentDate">
                            <small>Date of most recent payment</small>
                        </div>
                        <div class="form-group">
                            <label>Payment Status:</label>
                            <select id="paymentStatus">
                                <option value="Current">Current</option>
                                <option value="Late">Late</option>
                                <option value="Overdue">Overdue</option>
                                <option value="Partial">Partial</option>
                            </select>
                        </div>
                    </div>
                </div>
                
                <div class="form-section">
                    <div class="section-title">Additional Information</div>
                    <div class="form-group">
                        <label>Notes:</label>
                        <textarea id="notes" placeholder="Any additional notes about the tenant, lease terms, special arrangements, etc."></textarea>
                    </div>
                </div>
                
                <div style="margin-top: 30px; text-align: center;">
                    <button type="button" class="btn" onclick="submitManualApplication()">📝 Add Tenant to System</button>
                    <button type="button" class="btn btn-secondary" onclick="resetForm()">🔄 Clear Form</button>
                </div>
            </form>
        </div>
    `}
    
    <script>
        const availableRooms = ${roomsJson};
        
        // Set default dates
        document.addEventListener('DOMContentLoaded', function() {
            const today = new Date();
            const oneYearLater = new Date(today);
            oneYearLater.setFullYear(oneYearLater.getFullYear() + 1);
            
            document.getElementById('moveInDate').value = today.toISOString().split('T')[0];
            if (document.getElementById('leaseEndDate')) {
                document.getElementById('leaseEndDate').value = oneYearLater.toISOString().split('T')[0];
            }
        });
        
        function updateRoomInfo() {
            const roomSelect = document.getElementById('roomNumber');
            const selectedOption = roomSelect.options[roomSelect.selectedIndex];
            const roomInfo = document.getElementById('roomInfo');
            
            if (selectedOption.value) {
                const price = selectedOption.dataset.price;
                document.getElementById('rentalPriceDisplay').textContent = price;
                
                // Auto-populate security deposit with room price
                const priceNumber = parseFloat(price.replace('$', '').replace(',', ''));
                if (!isNaN(priceNumber)) {
                    document.getElementById('securityDeposit').value = priceNumber.toFixed(0);
                }
                
                roomInfo.style.display = 'block';
            } else {
                roomInfo.style.display = 'none';
            }
        }
        
        function submitManualApplication() {
            // Validate required fields
            const requiredFields = {
                'tenantName': 'Tenant Name',
                'tenantEmail': 'Tenant Email',
                'tenantPhone': 'Tenant Phone',
                'roomNumber': 'Room Number',
                'moveInDate': 'Move-In Date',
                'securityDeposit': 'Security Deposit'
            };
            
            let missingFields = [];
            for (let fieldId in requiredFields) {
                const field = document.getElementById(fieldId);
                if (!field.value || field.value.trim() === '') {
                    missingFields.push(requiredFields[fieldId]);
                }
            }
            
            if (missingFields.length > 0) {
                showStatus('Please fill in the following required fields: ' + missingFields.join(', '), 'error');
                return;
            }
            
            // Relaxed email validation - just check for @ symbol
            const email = document.getElementById('tenantEmail').value;
            if (!email.includes('@')) {
                showStatus('Please enter a valid email address (must contain @).', 'error');
                return;
            }
            
            // Validate move-in date is not too far in the past
            const moveInDate = new Date(document.getElementById('moveInDate').value);
            const today = new Date();
            const thirtyDaysAgo = new Date(today);
            thirtyDaysAgo.setDate(today.getDate() - 30);
            
            if (moveInDate < thirtyDaysAgo) {
                showStatus('Move-in date cannot be more than 30 days in the past.', 'error');
                return;
            }
            
            // Get room information
            const roomSelect = document.getElementById('roomNumber');
            const selectedRoom = availableRooms.find(room => room.number === roomSelect.value);
            
            const applicationData = {
                tenantName: document.getElementById('tenantName').value,
                tenantEmail: email,
                tenantPhone: document.getElementById('tenantPhone').value,
                emergencyContact: document.getElementById('emergencyContact').value,
                roomNumber: roomSelect.value,
                rentalPrice: selectedRoom.price,
                moveInDate: document.getElementById('moveInDate').value,
                leaseEndDate: document.getElementById('leaseEndDate').value,
                moveOutPlanned: document.getElementById('moveOutPlanned').value,
                negotiatedPrice: document.getElementById('negotiatedPrice').value,
                securityDeposit: document.getElementById('securityDeposit').value,
                lastPaymentDate: document.getElementById('lastPaymentDate').value,
                paymentStatus: document.getElementById('paymentStatus').value,
                notes: document.getElementById('notes').value
            };
            
            showStatus('Adding tenant to system...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    resetForm();
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .processManualApplication(JSON.stringify(applicationData));
        }
        
        function resetForm() {
            document.getElementById('manualApplicationForm').reset();
            
            // Reset default dates
            const today = new Date();
            document.getElementById('moveInDate').value = today.toISOString().split('T')[0];
            document.getElementById('paymentStatus').value = 'Current';
            
            // Hide room info
            document.getElementById('roomInfo').style.display = 'none';
            
            showStatus('Form cleared and ready for new tenant.', 'success');
        }
        
        function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = 'status ' + type;
            status.style.display = 'block';
            
            setTimeout(() => {
                if (type === 'success' && (message.includes('completed') || message.includes('cleared'))) {
                    status.style.display = 'none';
                }
            }, 5000);
        }
    </script>
</body>
</html>
    `;
  }
};

/**
 * Wrapper functions for menu integration
 */
function showManualApplicationPanel() {
  return ManualTenantApplication.showManualApplicationPanel();
}

function processManualApplication(applicationData) {
  return ManualTenantApplication.processManualApplication(applicationData);
}
