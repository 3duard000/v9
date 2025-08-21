/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Booking Manager - BookingManager.gs
 * 
 * Room Status Dashboard - Shows current room occupancy and upcoming bookings
 * without requiring date input. Displays real-time room status.
 */

const BookingManager = {

  /**
   * Show the main booking manager panel with room status dashboard
   */
  showBookingManagerPanel() {
    try {
      console.log('Opening Room Status Dashboard...');
      
      const html = this._generateRoomStatusDashboardHTML();
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(1200)
        .setHeight(800)
        .setTitle('🏨 Room Status Dashboard');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Room Status Dashboard');
      
    } catch (error) {
      console.error('Error showing room status dashboard:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load room status dashboard: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Get current room status for all rooms
   */
  getCurrentRoomStatus() {
    try {
      console.log('Getting current room status...');
      
      const rooms = this._getAllRoomsWithStatus();
      const summary = this._getRoomSummary(rooms);
      
      return JSON.stringify({
        success: true,
        rooms: rooms,
        summary: summary,
        lastUpdated: new Date().toLocaleString()
      });
      
    } catch (error) {
      console.error('Error getting room status:', error);
      return JSON.stringify({
        success: false,
        error: error.message
      });
    }
  },

  /**
   * Get all rooms with their current status and occupancy details
   * @private
   */
  _getAllRoomsWithStatus() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      let rooms = [];
      const today = new Date();
      
      if (!guestSheet || guestSheet.getLastRow() <= 1) {
        console.log('No guest data found, returning default rooms');
        return this._getDefaultRoomsWithStatus();
      }
      
      const data = guestSheet.getDataRange().getValues();
      const headers = data[0];
      
      // Find column indices
      const roomNumberCol = headers.indexOf('Room Number');
      const roomNameCol = headers.indexOf('Room Name');
      const roomTypeCol = headers.indexOf('Room Type');
      const dailyRateCol = headers.indexOf('Daily Rate');
      const checkInCol = headers.indexOf('Check-In Date');
      const checkOutCol = headers.indexOf('Check-Out Date');
      const currentGuestCol = headers.indexOf('Current Guest');
      const bookingStatusCol = headers.indexOf('Booking Status');
      const guestEmailCol = headers.indexOf('Guest Email');
      const guestPhoneCol = headers.indexOf('Guest Phone');
      
      if (roomNumberCol === -1) {
        return this._getDefaultRoomsWithStatus();
      }
      
      // Group data by room number
      const roomMap = new Map();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const roomNumber = row[roomNumberCol];
        
        if (!roomNumber) continue;
        
        if (!roomMap.has(roomNumber)) {
          roomMap.set(roomNumber, {
            roomNumber: roomNumber.toString(),
            roomName: (row[roomNameCol] || `Room ${roomNumber}`).toString(),
            roomType: (row[roomTypeCol] || 'Standard').toString(),
            dailyRate: (row[dailyRateCol] || '$75').toString(),
            bookings: []
          });
        }
        
        // Add booking if it has dates
        if (row[checkInCol] && row[checkOutCol]) {
          const checkIn = new Date(row[checkInCol]);
          const checkOut = new Date(row[checkOutCol]);
          const bookingStatus = row[bookingStatusCol] || '';
          
          roomMap.get(roomNumber).bookings.push({
            checkIn: checkIn,
            checkOut: checkOut,
            guest: row[currentGuestCol] || '',
            status: bookingStatus,
            email: row[guestEmailCol] || '',
            phone: row[guestPhoneCol] || ''
          });
        }
      }
      
      // Determine status for each room
      roomMap.forEach((room, roomNumber) => {
        const status = this._determineRoomStatus(room.bookings, today);
        rooms.push({
          roomNumber: room.roomNumber,
          roomName: room.roomName,
          roomType: room.roomType,
          dailyRate: room.dailyRate,
          status: status.status,
          statusColor: status.color,
          statusIcon: status.icon,
          currentGuest: status.currentGuest,
          checkIn: status.checkIn,
          checkOut: status.checkOut,
          nextBooking: status.nextBooking,
          bookings: room.bookings.filter(b => b.status !== 'Cancelled' && b.status !== 'Checked-Out')
        });
      });
      
      // Sort by room number
      rooms.sort((a, b) => a.roomNumber.localeCompare(b.roomNumber));
      
      console.log(`Found ${rooms.length} rooms with status`);
      return rooms;
      
    } catch (error) {
      console.error('Error getting rooms with status:', error);
      return this._getDefaultRoomsWithStatus();
    }
  },

  /**
   * Determine the current status of a room based on its bookings
   * @private
   */
  _determineRoomStatus(bookings, today) {
    const activeBookings = bookings.filter(b => 
      b.status !== 'Cancelled' && b.status !== 'Checked-Out'
    );
    
    // Check for current occupancy
    for (let booking of activeBookings) {
      if (today >= booking.checkIn && today <= booking.checkOut) {
        return {
          status: 'Occupied',
          color: '#ffebee',
          icon: '🔴',
          currentGuest: booking.guest,
          checkIn: booking.checkIn.toLocaleDateString(),
          checkOut: booking.checkOut.toLocaleDateString(),
          nextBooking: null
        };
      }
    }
    
    // Check for upcoming bookings (next 7 days)
    const upcomingBookings = activeBookings
      .filter(b => b.checkIn > today)
      .sort((a, b) => a.checkIn - b.checkIn);
    
    if (upcomingBookings.length > 0) {
      const nextBooking = upcomingBookings[0];
      const daysUntilBooking = Math.ceil((nextBooking.checkIn - today) / (1000 * 60 * 60 * 24));
      
      return {
        status: 'Available',
        color: '#e8f5e8',
        icon: '🟢',
        currentGuest: '',
        checkIn: '',
        checkOut: '',
        nextBooking: {
          guest: nextBooking.guest,
          checkIn: nextBooking.checkIn.toLocaleDateString(),
          daysUntil: daysUntilBooking
        }
      };
    }
    
    // Room is available with no upcoming bookings
    return {
      status: 'Available',
      color: '#e8f5e8',
      icon: '🟢',
      currentGuest: '',
      checkIn: '',
      checkOut: '',
      nextBooking: null
    };
  },

  /**
   * Get default rooms when no data exists
   * @private
   */
  _getDefaultRoomsWithStatus() {
    return [
      {
        roomNumber: '201',
        roomName: 'Garden View Suite',
        roomType: 'Deluxe',
        dailyRate: '$85',
        status: 'Available',
        statusColor: '#e8f5e8',
        statusIcon: '🟢',
        currentGuest: '',
        checkIn: '',
        checkOut: '',
        nextBooking: null,
        bookings: []
      },
      {
        roomNumber: '202',
        roomName: 'City View Room',
        roomType: 'Standard',
        dailyRate: '$65',
        status: 'Available',
        statusColor: '#e8f5e8',
        statusIcon: '🟢',
        currentGuest: '',
        checkIn: '',
        checkOut: '',
        nextBooking: null,
        bookings: []
      },
      {
        roomNumber: '203',
        roomName: 'Executive Suite',
        roomType: 'Premium',
        dailyRate: '$120',
        status: 'Available',
        statusColor: '#e8f5e8',
        statusIcon: '🟢',
        currentGuest: '',
        checkIn: '',
        checkOut: '',
        nextBooking: null,
        bookings: []
      },
      {
        roomNumber: '204',
        roomName: 'Economy Room',
        roomType: 'Standard',
        dailyRate: '$55',
        status: 'Available',
        statusColor: '#e8f5e8',
        statusIcon: '🟢',
        currentGuest: '',
        checkIn: '',
        checkOut: '',
        nextBooking: null,
        bookings: []
      }
    ];
  },

  /**
   * Get summary statistics
   * @private
   */
  _getRoomSummary(rooms) {
    const total = rooms.length;
    const occupied = rooms.filter(r => r.status === 'Occupied').length;
    const available = rooms.filter(r => r.status === 'Available').length;
    const occupancyRate = total > 0 ? Math.round((occupied / total) * 100) : 0;
    
    return {
      total: total,
      occupied: occupied,
      available: available,
      occupancyRate: occupancyRate
    };
  },

  /**
   * Create new booking
   */
  createNewBooking(bookingData) {
    try {
      console.log('Creating new booking:', bookingData);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      if (!guestSheet) {
        throw new Error('Guest Rooms sheet not found');
      }
      
      const data = JSON.parse(bookingData);
      
      // Generate booking ID and calculate totals
      const bookingId = this._generateBookingId();
      const nights = this._calculateNights(new Date(data.checkInDate), new Date(data.checkOutDate));
      const totalAmount = parseFloat(data.dailyRate || 75) * nights;
      
      // Create booking row
      const bookingRow = [
        bookingId,                           // Booking ID
        data.roomNumber,                     // Room Number
        data.roomName || 'Guest Room',      // Room Name
        data.roomType || 'Standard',        // Room Type
        `$${parseFloat(data.dailyRate || 75).toFixed(2)}`, // Daily Rate
        data.checkInDate,                    // Check-In Date
        data.checkOutDate,                   // Check-Out Date
        nights.toString(),                   // Number of Nights
        data.numberOfGuests || '1',          // Number of Guests
        data.guestName,                      // Current Guest
        data.guestEmail || '',               // Guest Email
        data.guestPhone || '',               // Guest Phone
        data.purposeOfVisit || '',           // Purpose of Visit
        `$${totalAmount.toFixed(2)}`,       // Total Amount
        data.paymentStatus || 'Pending',     // Payment Status
        'Reserved',                          // Booking Status
        data.bookingSource || 'Direct',      // Source
        `Booking created on ${new Date().toLocaleDateString()}` // Notes
      ];
      
      // Add to guest rooms sheet
      const lastRow = guestSheet.getLastRow();
      guestSheet.getRange(lastRow + 1, 1, 1, bookingRow.length).setValues([bookingRow]);
      
      console.log(`Booking ${bookingId} created successfully`);
      return `✅ Booking ${bookingId} created successfully for ${data.guestName} in Room ${data.roomNumber}`;
      
    } catch (error) {
      console.error('Error creating booking:', error);
      throw new Error('Failed to create booking: ' + error.message);
    }
  },

  /**
   * Generate unique booking ID
   * @private
   */
  _generateBookingId() {
    const prefix = 'BK';
    const timestamp = Date.now().toString().slice(-6);
    const random = Math.floor(Math.random() * 100).toString().padStart(2, '0');
    return `${prefix}${timestamp}${random}`;
  },

  /**
   * Calculate number of nights
   * @private
   */
  _calculateNights(startDate, endDate) {
    const timeDiff = endDate.getTime() - startDate.getTime();
    return Math.ceil(timeDiff / (1000 * 3600 * 24));
  },

  /**
   * Generate HTML for room status dashboard
   * @private
   */
  _generateRoomStatusDashboardHTML() {
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
            margin: 0; 
            background: #f8f9fa;
        }
        .header { 
            background: linear-gradient(135deg, #1c4587 0%, #2d5aa0 100%); 
            color: white; 
            padding: 20px; 
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .header h1 { margin: 0; font-size: 24px; }
        .header p { margin: 5px 0 0 0; opacity: 0.9; }
        
        .dashboard-container { padding: 20px; max-width: 1400px; margin: 0 auto; }
        
        .summary-cards { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); 
            gap: 15px; 
            margin-bottom: 30px;
        }
        
        .summary-card {
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-left: 4px solid #1c4587;
        }
        
        .summary-card h3 { margin: 0 0 10px 0; color: #666; font-size: 14px; text-transform: uppercase; }
        .summary-card .value { font-size: 28px; font-weight: bold; margin: 0; color: #1c4587; }
        .summary-card .subtitle { font-size: 12px; color: #999; margin-top: 5px; }
        
        .rooms-grid { 
            display: grid; 
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr)); 
            gap: 20px; 
        }
        
        .room-card {
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-left: 6px solid #ddd;
            transition: all 0.3s ease;
        }
        
        .room-card:hover { transform: translateY(-2px); box-shadow: 0 4px 16px rgba(0,0,0,0.15); }
        .room-available { border-left-color: #22c55e; }
        .room-occupied { border-left-color: #ef4444; }
        .room-maintenance { border-left-color: #f59e0b; }
        
        .room-header {
            display: flex;
            justify-content: between;
            align-items: center;
            margin-bottom: 15px;
        }
        
        .room-title { 
            font-size: 18px; 
            font-weight: bold; 
            margin: 0;
            color: #1f2937;
        }
        
        .room-status {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
        }
        
        .status-available { background: #dcfce7; color: #16a34a; }
        .status-occupied { background: #fee2e2; color: #dc2626; }
        .status-maintenance { background: #fef3c7; color: #d97706; }
        
        .room-details { margin-bottom: 15px; }
        .room-detail { 
            display: flex; 
            justify-content: space-between; 
            margin: 5px 0; 
            font-size: 14px;
        }
        .room-detail strong { color: #374151; }
        
        .current-guest {
            background: #f3f4f6;
            border-radius: 8px;
            padding: 12px;
            margin-top: 10px;
        }
        
        .current-guest h4 { margin: 0 0 8px 0; color: #1f2937; font-size: 14px; }
        .current-guest p { margin: 2px 0; font-size: 13px; color: #6b7280; }
        
        .next-booking {
            background: #fffbeb;
            border: 1px solid #fed7aa;
            border-radius: 8px;
            padding: 12px;
            margin-top: 10px;
        }
        
        .next-booking h4 { margin: 0 0 8px 0; color: #92400e; font-size: 14px; }
        .next-booking p { margin: 2px 0; font-size: 13px; color: #a3a3a3; }
        
        .refresh-btn {
            position: fixed;
            bottom: 30px;
            right: 30px;
            background: #1c4587;
            color: white;
            border: none;
            border-radius: 50px;
            padding: 15px 25px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            box-shadow: 0 4px 12px rgba(28, 69, 135, 0.3);
            transition: all 0.3s ease;
        }
        
        .refresh-btn:hover { 
            background: #1e40af; 
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(28, 69, 135, 0.4);
        }
        
        .loading { text-align: center; padding: 40px; color: #6b7280; }
        .error { text-align: center; padding: 40px; color: #dc2626; }
        
        .last-updated {
            text-align: center;
            margin-top: 20px;
            font-size: 12px;
            color: #9ca3af;
        }
        
        @media (max-width: 768px) {
            .dashboard-container { padding: 15px; }
            .rooms-grid { grid-template-columns: 1fr; }
            .summary-cards { grid-template-columns: repeat(2, 1fr); }
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>🏨 Room Status Dashboard</h1>
        <p>Real-time view of all room occupancy and upcoming bookings</p>
    </div>
    
    <div class="dashboard-container">
        <div id="summary-section" class="summary-cards">
            <!-- Summary cards will be populated here -->
        </div>
        
        <div id="rooms-section" class="rooms-grid">
            <div class="loading">
                <p>🔄 Loading room status...</p>
            </div>
        </div>
        
        <div id="last-updated" class="last-updated"></div>
    </div>
    
    <button class="refresh-btn" onclick="loadRoomStatus()">
        🔄 Refresh
    </button>
    
    <script>
        // Load room status when page loads
        document.addEventListener('DOMContentLoaded', function() {
            loadRoomStatus();
            
            // Auto-refresh every 2 minutes
            setInterval(loadRoomStatus, 120000);
        });
        
        function loadRoomStatus() {
            console.log('Loading room status...');
            
            // Show loading state
            document.getElementById('rooms-section').innerHTML = '<div class="loading"><p>🔄 Loading room status...</p></div>';
            
            google.script.run
                .withSuccessHandler(function(result) {
                    console.log('Room status loaded:', result);
                    try {
                        const data = JSON.parse(result);
                        if (data.success) {
                            displayRoomStatus(data);
                        } else {
                            showError(data.error || 'Unknown error occurred');
                        }
                    } catch (parseError) {
                        console.error('Parse error:', parseError);
                        showError('Failed to parse room status data');
                    }
                })
                .withFailureHandler(function(error) {
                    console.error('Failed to load room status:', error);
                    showError('Failed to load room status: ' + error.message);
                })
                .getCurrentRoomStatus();
        }
        
        function displayRoomStatus(data) {
            displaySummary(data.summary);
            displayRooms(data.rooms);
            document.getElementById('last-updated').textContent = 'Last updated: ' + data.lastUpdated;
        }
        
        function displaySummary(summary) {
            const summarySection = document.getElementById('summary-section');
            
            summarySection.innerHTML = \`
                <div class="summary-card">
                    <h3>Total Rooms</h3>
                    <p class="value">\${summary.total}</p>
                    <p class="subtitle">Active guest rooms</p>
                </div>
                <div class="summary-card">
                    <h3>Occupied</h3>
                    <p class="value" style="color: #ef4444">\${summary.occupied}</p>
                    <p class="subtitle">Currently occupied</p>
                </div>
                <div class="summary-card">
                    <h3>Available</h3>
                    <p class="value" style="color: #22c55e">\${summary.available}</p>
                    <p class="subtitle">Ready for booking</p>
                </div>
                <div class="summary-card">
                    <h3>Occupancy Rate</h3>
                    <p class="value" style="color: #1c4587">\${summary.occupancyRate}%</p>
                    <p class="subtitle">Current utilization</p>
                </div>
            \`;
        }
        
        function displayRooms(rooms) {
            const roomsSection = document.getElementById('rooms-section');
            
            if (rooms.length === 0) {
                roomsSection.innerHTML = '<div class="error"><p>❌ No rooms found. Please add rooms to the Guest Rooms sheet.</p></div>';
                return;
            }
            
            let html = '';
            
            rooms.forEach(room => {
                const statusClass = 'status-' + room.status.toLowerCase();
                const cardClass = 'room-' + room.status.toLowerCase();
                
                html += \`
                    <div class="room-card \${cardClass}">
                        <div class="room-header">
                            <h3 class="room-title">\${room.statusIcon} Room \${room.roomNumber}</h3>
                            <span class="room-status \${statusClass}">\${room.status}</span>
                        </div>
                        
                        <div class="room-details">
                            <div class="room-detail">
                                <span><strong>Name:</strong></span>
                                <span>\${room.roomName}</span>
                            </div>
                            <div class="room-detail">
                                <span><strong>Type:</strong></span>
                                <span>\${room.roomType}</span>
                            </div>
                            <div class="room-detail">
                                <span><strong>Rate:</strong></span>
                                <span>\${room.dailyRate}/night</span>
                            </div>
                        </div>
                \`;
                
                // Current guest info
                if (room.currentGuest && room.status === 'Occupied') {
                    html += \`
                        <div class="current-guest">
                            <h4>👤 Current Guest</h4>
                            <p><strong>\${room.currentGuest}</strong></p>
                            <p>Check-in: \${room.checkIn}</p>
                            <p>Check-out: \${room.checkOut}</p>
                        </div>
                    \`;
                }
                
                // Next booking info
                if (room.nextBooking && room.status === 'Available') {
                    html += \`
                        <div class="next-booking">
                            <h4>📅 Upcoming Booking</h4>
                            <p><strong>\${room.nextBooking.guest}</strong></p>
                            <p>Arrives: \${room.nextBooking.checkIn}</p>
                            <p>In \${room.nextBooking.daysUntil} day(s)</p>
                        </div>
                    \`;
                }
                
                html += '</div>';
            });
            
            roomsSection.innerHTML = html;
        }
        
        function showError(message) {
            document.getElementById('rooms-section').innerHTML = \`
                <div class="error">
                    <p>❌ \${message}</p>
                    <button onclick="loadRoomStatus()" style="margin-top: 10px; padding: 8px 16px; border: none; background: #1c4587; color: white; border-radius: 4px; cursor: pointer;">Try Again</button>
                </div>
            \`;
        }
    </script>
</body>
</html>
    `;
  }
};
