/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Main Entry Point - Main.gs
 * 
 * UPDATED: Added Manual Application Entry to menu while keeping existing Google Form processing
 */

/**
 * Create custom menu when spreadsheet opens
 * UPDATED: Added Quick Actions submenu for common operations
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🏠 White House Manager')
    //.addItem('⚙️ Initialize System', 'setupTenantManagement')
     .addSeparator()
    .addSubMenu(ui.createMenu('📋 Tenant Management')
      .addItem('📝 Manual Application Entry', 'showManualApplicationPanel')
      .addItem('💰 Record Payments', 'showPaymentRecordingPanel')
      .addItem('📋 Process Online Applications', 'showApplicationProcessingPanel')
      .addItem('🚪 Process Move-Outs', 'showMoveOutProcessingPanel'))
       .addSeparator()
    .addSubMenu(ui.createMenu('🏨 Guest Management')
      .addItem('📊 Room Status Dashboard', 'showBookingManagerPanel')
      .addItem('📝 Create New Booking', 'showNewBookingPanel')
      .addItem('🏨 Process Online Reservations', 'showOnlineReservationPanel')
      .addItem('✅ Process Check-In', 'showCheckInPanel')
      .addItem('🚪 Process Check-Out', 'showCheckOutPanel'))
       .addSeparator()
    .addSubMenu(ui.createMenu('📧 Emails')
      .addItem('💸 Send Rent Reminders', 'sendRentReminders')
      .addItem('⚠️ Send Late Payment Alerts', 'sendLatePaymentAlerts')
      .addItem('📋 Send Monthly Invoices', 'sendMonthlyInvoices'))
      //.addItem('🔄 Update Payment Status', 'updateAllPaymentStatuses'))
       .addSeparator()
    .addSubMenu(ui.createMenu('⚡ Quick Actions')
      .addItem('💵 Add Income/Expenses', 'showBudgetEntryPanel')
      .addItem('🔧 Add Maintenance Request', 'showMaintenanceRequestPanel')
      //.addItem('🔄 Rename Form Sheets', 'performDelayedSheetRenaming')
      .addItem('🔄 Refresh All Dashboards', 'refreshAllDashboards'))
       .addSeparator()
    .addToUi();
}

/**
 * Main setup function - run this first to initialize everything
 */
function setupTenantManagement() {
  try {
    console.log('Setting up Tenant Management System...');
    
    // Create sheets with headers and formatting
    SheetManager.createRequiredSheets();
    
    // Add sample data to demonstrate the system
    DataManager.addSampleData();
    
    // Create and link Google Forms
    FormManager.createGoogleForms();
    
    // Set up automated email triggers
    TriggerManager.setupTriggers();
    
    // Create initial dashboards
    Dashboard.createManagementDashboard();
    Dashboard.createFinancialDashboard();
    
    // Set up delayed sheet renaming after forms are created
    console.log('Setting up delayed sheet renaming...');
    FormManager.setupDelayedSheetRenaming();
    
    console.log('Setup completed successfully!');
    return 'Tenant Management System setup completed successfully! ' +
           'Dashboards created and will auto-refresh 3x daily. ' +
           'Form response sheets will be automatically renamed in 3 minutes. ' +
           'Check the execution log for form URLs.';
  } catch (error) {
    console.error('Setup failed:', error);
    throw new Error('Setup failed: ' + error.message);
  }
}

/**
 * Manual Tenant Application wrapper functions
 */
function showManualApplicationPanel() {
  return ManualTenantApplication.showManualApplicationPanel();
}

function processManualApplication(applicationData) {
  return ManualTenantApplication.processManualApplication(applicationData);
}

/**
 * Room Status Dashboard (replaces old availability checker)
 */
function showBookingManagerPanel() {
  return BookingManager.showBookingManagerPanel();
}

/**
 * Function to get current room status for the dashboard
 */
function getCurrentRoomStatus() {
  return BookingManager.getCurrentRoomStatus();
}

/**
 * Individual Booking Panel wrapper functions
 */
function showNewBookingPanel() {
  return BookingPanels.showNewBookingPanel();
}

function showCheckInPanel() {
  return BookingPanels.showCheckInPanel();
}

function showCheckOutPanel() {
  return BookingPanels.showCheckOutPanel();
}

function showOnlineReservationPanel() {
  return BookingPanels.showOnlineReservationPanel();
}

function processOnlineReservation(reservationData) {
  return BookingPanels.processOnlineReservation(reservationData);
}

/**
 * OLD FUNCTIONS - Keep these for backward compatibility but they're no longer in the menu
 */
function showAvailabilityChecker() {
  // Redirect to new Room Status Dashboard
  return BookingManager.showBookingManagerPanel();
}

function checkAvailability(availabilityData) {
  // Keep for any legacy calls
  return BookingManager.checkAvailability(availabilityData);
}

/**
 * Booking creation and management functions
 */
function createNewBooking(bookingData) {
  return BookingManager.createNewBooking(bookingData);
}

function processCheckIn(checkInData) {
  return BookingManager.processCheckIn(checkInData);
}

function processCheckOut(checkOutData) {
  return BookingManager.processCheckOut(checkOutData);
}

function addTenantToCalendar(tenantData) {
  return BookingManager.addTenantToCalendar(tenantData);
}

/**
 * Calendar integration wrapper functions
 */
function syncAllTenantsToCalendar() {
  return CalendarManager.syncAllTenantsToCalendar();
}

function syncAllGuestsToCalendar() {
  return CalendarManager.syncAllGuestsToCalendar();
}

function syncAllToCalendar() {
  return CalendarManager.syncAllToCalendar();
}

function clearWhiteHouseCalendarEvents() {
  return CalendarManager.clearWhiteHouseCalendarEvents();
}

/**
 * Application Processing Panel wrapper function
 */
function showApplicationProcessingPanel() {
  return Panel.showApplicationProcessingPanel();
}

function processApplication(applicationData) {
  return Panel.processApplication(applicationData);
}

function markApplicationAsRejected(applicationData) {
  return Panel.markApplicationAsRejected(applicationData);
}

/**
 * Move-Out Processing Panel wrapper function
 */
function showMoveOutProcessingPanel() {
  return MoveOutPanel.showMoveOutProcessingPanel();
}

function processMoveOutRequest(moveOutData) {
  return MoveOutPanel.processMoveOutRequest(moveOutData);
}

/**
 * Payment Recording Panel wrapper function
 */
function showPaymentRecordingPanel() {
  return PaymentPanel.showPaymentRecordingPanel();
}

function recordTenantPayment(paymentData) {
  return PaymentPanel.recordTenantPayment(paymentData);
}

/**
 * Budget Entry Panel wrapper function
 */
function showBudgetEntryPanel() {
  return BudgetPanel.showBudgetEntryPanel();
}

function addBudgetEntry(entryData) {
  return BudgetPanel.addBudgetEntry(entryData);
}

/**
 * Maintenance Request Panel wrapper function
 */
function showMaintenanceRequestPanel() {
  return MaintenancePanel.showMaintenanceRequestPanel();
}

function addMaintenanceRequest(requestData) {
  return MaintenancePanel.addMaintenanceRequest(requestData);
}

/**
 * Dashboard wrapper functions
 */
function createManagementDashboard() {
  return Dashboard.createManagementDashboard();
}

function createFinancialDashboard() {
  return Dashboard.createFinancialDashboard();
}

function refreshAllDashboards() {
  return Dashboard.refreshAllDashboards();
}

/**
 * Email wrapper functions
 */
function sendRentReminders() {
  return EmailManager.sendRentReminders();
}

function sendLatePaymentAlerts() {
  return EmailManager.sendLatePaymentAlerts();
}

function sendMonthlyInvoices() {
  return EmailManager.sendMonthlyInvoices();
}

function updateAllPaymentStatuses() {
  return EmailManager.updateAllPaymentStatuses();
}

function sendAllPaymentAlerts() {
  return EmailManager.sendLatePaymentAlerts();
}

/**
 * Other wrapper functions
 */
function performDelayedSheetRenaming() {
  try {
    console.log('Manually triggering sheet renaming...');
    return FormManager.performDelayedSheetRenaming();
  } catch (error) {
    console.error('Manual sheet renaming failed:', error);
    throw new Error('Sheet renaming failed: ' + error.message);
  }
}

/**
 * Daily check function that runs the appropriate email functions based on the date
 */
function checkAndRunDailyTasks() {
  const today = new Date();
  const dayOfMonth = today.getDate();
  const currentHour = today.getHours();
  
  console.log(`Daily check running - Day: ${dayOfMonth}, Hour: ${currentHour}`);
  
  // Only run email functions at 9 AM to avoid multiple executions
  if (currentHour === 9) {
    
    // Update payment statuses every day (this runs first)
    console.log('Updating all tenant payment statuses...');
    EmailManager.updateAllPaymentStatuses();
    
    // Sync calendar data daily at 9 AM
    console.log('Daily calendar sync...');
    try {
      CalendarManager.syncAllTenantsToCalendar();
      CalendarManager.syncAllGuestsToCalendar();
      console.log('✅ Daily calendar sync completed');
    } catch (calendarError) {
      console.log('⚠️ Daily calendar sync failed:', calendarError.message);
    }
    
    // Send rent reminders one week before due date (25th of each month)
    if (dayOfMonth === 25) {
      console.log('Running rent reminders (one week before due date)...');
      EmailManager.sendRentReminders();
    }
    
    // Send monthly invoices on the 1st (due date)
    if (dayOfMonth === 1) {
      console.log('Running monthly invoices (due date)...');
      EmailManager.sendMonthlyInvoices();
    }
    
    // Send late payment alerts starting the day after due date (2nd) and every week
    if (dayOfMonth === 2 || dayOfMonth === 9 || dayOfMonth === 16 || dayOfMonth === 23) {
      console.log(`Running late payment alerts (day ${dayOfMonth})...`);
      EmailManager.sendLatePaymentAlerts();
    }
  }
}

/**
 * Utility functions for testing and debugging
 */
function testRentReminder() {
  return EmailManager.testRentReminder();
}

function getTriggerInfo() {
  return TriggerManager.getTriggerInfo();
}
