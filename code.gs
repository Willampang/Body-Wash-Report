function doGet() {
  return HtmlService.createHtmlOutputFromFile('form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// NEW: Get totals up to yesterday (excluding today's data if any)
function getCurrentTotalsForPreview() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = spreadsheet.getSheetByName("Sheet1");
  
  if (!summarySheet) {
    // If no data exists, return original base values
    return {
      fb: { enquiry: 1185, followup: 649, waiting: 8, noreply: 507, drop: 7, closedCustomers: 33, closedB3F1: 28, closedSingle: 6 },
      organic: { enquiry: 5, followup: 0, waiting: 0, noreply: 0, drop: 0, closedCustomers: 5, closedB3F1: 5, closedSingle: 0 }
    };
  }
  
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var lastRow = summarySheet.getLastRow();
  
  if (lastRow <= 1) {
    // No data entries, return original base values
    return {
      fb: { enquiry: 1185, followup: 649, waiting: 8, noreply: 507, drop: 7, closedCustomers: 33, closedB3F1: 28, closedSingle: 6 },
      organic: { enquiry: 5, followup: 0, waiting: 0, noreply: 0, drop: 0, closedCustomers: 5, closedB3F1: 5, closedSingle: 0 }
    };
  }
  
  // Get all data
  var allData = summarySheet.getRange(2, 1, lastRow - 1, 15).getValues();
  
  // Filter out today's entries and calculate totals up to yesterday
  var totals = {
    fb: { enquiry: 1185, followup: 649, waiting: 8, noreply: 507, drop: 7, closedCustomers: 33, closedB3F1: 28, closedSingle: 6 },
    organic: { enquiry: 5, followup: 0, waiting: 0, noreply: 0, drop: 0, closedCustomers: 5, closedB3F1: 5, closedSingle: 0 }
  };
  
  // Add up all entries that are NOT from today
  for (var i = 0; i < allData.length; i++) {
    var entryDate = allData[i][0]; // Date column
    if (entryDate !== today) { // Only include entries that are not from today
      // FB data (columns B-G)
      totals.fb.enquiry += allData[i][1] || 0;
      totals.fb.followup += allData[i][2] || 0;
      totals.fb.waiting += allData[i][3] || 0;
      totals.fb.noreply += allData[i][4] || 0;
      totals.fb.drop += allData[i][5] || 0;
      
      var fbClosedProducts = allData[i][6] || 0;
      // Estimate customers and product breakdown (you may need to adjust these ratios)
      totals.fb.closedCustomers += Math.floor(fbClosedProducts * 0.9); // Assume 90% of products = customers
      totals.fb.closedB3F1 += Math.floor(fbClosedProducts * 0.8); // Assume 80% are B3F1 sets
      totals.fb.closedSingle += Math.floor(fbClosedProducts * 0.2); // Assume 20% are singles
      
      // Organic data (columns H-M)
      totals.organic.enquiry += allData[i][7] || 0;
      totals.organic.followup += allData[i][8] || 0;
      totals.organic.waiting += allData[i][9] || 0;
      totals.organic.noreply += allData[i][10] || 0;
      totals.organic.drop += allData[i][11] || 0;
      
      var organicClosedProducts = allData[i][12] || 0;
      // Estimate customers and product breakdown
      totals.organic.closedCustomers += Math.floor(organicClosedProducts * 0.9);
      totals.organic.closedB3F1 += Math.floor(organicClosedProducts * 0.8);
      totals.organic.closedSingle += Math.floor(organicClosedProducts * 0.2);
    }
  }
  
  return totals;
}
// NEW: Get current base totals from a dedicated config sheet
function getBaseTotals() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = spreadsheet.getSheetByName("Config");
  
  if (!configSheet) {
    // Create config sheet with initial values
    configSheet = spreadsheet.insertSheet("Config");
    configSheet.appendRow(["Setting", "FB_Enquiry", "FB_Followup", "FB_Waiting", "FB_Noreply", "FB_Drop", "FB_ClosedCustomers", "FB_ClosedB3F1", "FB_ClosedSingle", "Organic_Enquiry", "Organic_Followup", "Organic_Waiting", "Organic_Noreply", "Organic_Drop", "Organic_ClosedCustomers", "Organic_ClosedB3F1", "Organic_ClosedSingle"]);
    configSheet.appendRow(["BaseTotals", 1185, 649, 8, 507, 7, 33, 28, 6, 5, 0, 0, 0, 0, 5, 5, 0]);
    return {
      fb: { enquiry: 1185, followup: 649, waiting: 8, noreply: 507, drop: 7, closedCustomers: 33, closedB3F1: 28, closedSingle: 6 },
      organic: { enquiry: 5, followup: 0, waiting: 0, noreply: 0, drop: 0, closedCustomers: 5, closedB3F1: 5, closedSingle: 0 }
    };
  }
  
  var data = configSheet.getRange(2, 2, 1, 16).getValues()[0];
  return {
    fb: {
      enquiry: data[0] || 1185,
      followup: data[1] || 649,
      waiting: data[2] || 8,
      noreply: data[3] || 507,
      drop: data[4] || 7,
      closedCustomers: data[5] || 33,
      closedB3F1: data[6] || 28,
      closedSingle: data[7] || 6
    },
    organic: {
      enquiry: data[8] || 5,
      followup: data[9] || 0,
      waiting: data[10] || 0,
      noreply: data[11] || 0,
      drop: data[12] || 0,
      closedCustomers: data[13] || 5,
      closedB3F1: data[14] || 5,
      closedSingle: data[15] || 0
    }
  };
}

// NEW: Update base totals in config sheet
function updateBaseTotals(newData) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = spreadsheet.getSheetByName("Config");
  
  if (!configSheet) {
    getBaseTotals(); // This will create the config sheet
    configSheet = spreadsheet.getSheetByName("Config");
  }
  
  var currentTotals = getBaseTotals();
  
  // Add new data to current totals
  currentTotals.fb.enquiry += parseInt(newData.fbEnquiry || 0);
  currentTotals.fb.followup += parseInt(newData.fbFollowup || 0);
  currentTotals.fb.waiting += parseInt(newData.fbWaiting || 0);
  currentTotals.fb.noreply += parseInt(newData.fbNoreply || 0);
  currentTotals.fb.drop += parseInt(newData.fbDrop || 0);
  currentTotals.fb.closedCustomers += parseInt(newData.fbClosedCustomers || 0);
  currentTotals.fb.closedB3F1 += parseInt(newData.fbClosedB3F1 || 0);
  currentTotals.fb.closedSingle += parseInt(newData.fbClosedSingle || 0);
  
  currentTotals.organic.enquiry += parseInt(newData.organicEnquiry || 0);
  currentTotals.organic.followup += parseInt(newData.organicFollowup || 0);
  currentTotals.organic.waiting += parseInt(newData.organicWaiting || 0);
  currentTotals.organic.noreply += parseInt(newData.organicNoreply || 0);
  currentTotals.organic.drop += parseInt(newData.organicDrop || 0);
  currentTotals.organic.closedCustomers += parseInt(newData.organicClosedCustomers || 0);
  currentTotals.organic.closedB3F1 += parseInt(newData.organicClosedB3F1 || 0);
  currentTotals.organic.closedSingle += parseInt(newData.organicClosedSingle || 0);
  
  // Save updated totals back to config sheet
  configSheet.getRange(2, 2, 1, 16).setValues([[
    currentTotals.fb.enquiry, currentTotals.fb.followup, currentTotals.fb.waiting, currentTotals.fb.noreply, currentTotals.fb.drop, currentTotals.fb.closedCustomers, currentTotals.fb.closedB3F1, currentTotals.fb.closedSingle,
    currentTotals.organic.enquiry, currentTotals.organic.followup, currentTotals.organic.waiting, currentTotals.organic.noreply, currentTotals.organic.drop, currentTotals.organic.closedCustomers, currentTotals.organic.closedB3F1, currentTotals.organic.closedSingle
  ]]);
  
  return currentTotals;
}

// NEW: Reset base totals to original values
function resetBaseTotals() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = spreadsheet.getSheetByName("Config");
  
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet("Config");
    configSheet.appendRow(["Setting", "FB_Enquiry", "FB_Followup", "FB_Waiting", "FB_Noreply", "FB_Drop", "FB_ClosedCustomers", "FB_ClosedB3F1", "FB_ClosedSingle", "Organic_Enquiry", "Organic_Followup", "Organic_Waiting", "Organic_Noreply", "Organic_Drop", "Organic_ClosedCustomers", "Organic_ClosedB3F1", "Organic_ClosedSingle"]);
  }
  
  // Reset to original values
  configSheet.getRange(2, 1, 1, 17).setValues([["BaseTotals", 1185, 649, 8, 507, 7, 33, 28, 6, 5, 0, 0, 0, 0, 5, 5, 0]]);
  
  return {
    fb: { enquiry: 1185, followup: 649, waiting: 8, noreply: 507, drop: 7, closedCustomers: 33, closedB3F1: 28, closedSingle: 6 },
    organic: { enquiry: 5, followup: 0, waiting: 0, noreply: 0, drop: 0, closedCustomers: 5, closedB3F1: 5, closedSingle: 0 }
  };
}

function submitData(data) {
  try {
    // Get current base totals
    var currentTotals = getBaseTotals();
    
    // Get the active spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Try to get existing Sheet1 or create it
    var summarySheet = spreadsheet.getSheetByName("Sheet1");
    if (!summarySheet) {
      summarySheet = spreadsheet.insertSheet("Sheet1");
      // Add headers to match your current spreadsheet structure
      summarySheet.appendRow([
        "Date", 
        "Enquiry", "Follow Up", "Waiting Payment", "No Reply", "Drop", "Closed",
        "Enquiry(Whatsapp)", "Follow Up(Whatsapp)", "Waiting Payment(Whatsapp)", "No Reply(Whatsapp)", "Drop(Whatsapp)", "Closed(Whatsapp)",
        "Total Sales", "Today's Closed"
      ]);
    }
    
    var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    // Convert FB Ad data to numbers (maps to regular columns)
    var fbEnquiry = parseInt(data.fbEnquiry) || 0;
    var fbFollowup = parseInt(data.fbFollowup) || 0;
    var fbWaiting = parseInt(data.fbWaiting) || 0;
    var fbNoreply = parseInt(data.fbNoreply) || 0;
    var fbDrop = parseInt(data.fbDrop) || 0;
    var fbClosedB3F1 = parseInt(data.fbClosedB3F1) || 0;
    var fbClosedSingle = parseInt(data.fbClosedSingle) || 0;
    var fbClosedTotal = fbClosedB3F1 + fbClosedSingle;
    var fbClosedCustomers = parseInt(data.fbClosedCustomers) || 0;
    
    // Convert Organic data to numbers (maps to Whatsapp columns)
    var organicEnquiry = parseInt(data.organicEnquiry) || 0;
    var organicFollowup = parseInt(data.organicFollowup) || 0;
    var organicWaiting = parseInt(data.organicWaiting) || 0;
    var organicNoreply = parseInt(data.organicNoreply) || 0;
    var organicDrop = parseInt(data.organicDrop) || 0;
    var organicClosedB3F1 = parseInt(data.organicClosedB3F1) || 0;
    var organicClosedSingle = parseInt(data.organicClosedSingle) || 0;
    var organicClosedTotal = organicClosedB3F1 + organicClosedSingle;
    var organicClosedCustomers = parseInt(data.organicClosedCustomers) || 0;
    
    // Calculate today's closed total from customer count fields
    var todaysClosed = fbClosedCustomers + organicClosedCustomers;
    
    // Calculate total sales (B3F1 * 487 + Single * 167)
    var totalSales = (fbClosedB3F1 * 487) + (fbClosedSingle * 167) + (organicClosedB3F1 * 487) + (organicClosedSingle * 167);
    
    // Save to summary sheet matching your column structure
    summarySheet.appendRow([
      date, 
      fbEnquiry, fbFollowup, fbWaiting, fbNoreply, fbDrop, fbClosedTotal,
      organicEnquiry, organicFollowup, organicWaiting, organicNoreply, organicDrop, organicClosedTotal,
      totalSales, todaysClosed
    ]);
    
    // Update base totals with submitted data
    var updatedTotals = updateBaseTotals(data);
    
    // Append new daily report to the Daily Report sheet
    appendToDailyReport(spreadsheet, {
      date: today,
      fb: {
        enquiry: fbEnquiry,
        followup: fbFollowup,
        waiting: fbWaiting,
        noreply: fbNoreply,
        drop: fbDrop,
        closedCustomers: fbClosedCustomers,
        closedB3F1: fbClosedB3F1,
        closedSingle: fbClosedSingle,
        closedTotal: fbClosedTotal
      },
      organic: {
        enquiry: organicEnquiry,
        followup: organicFollowup,
        waiting: organicWaiting,
        noreply: organicNoreply,
        drop: organicDrop,
        closedCustomers: organicClosedCustomers,
        closedB3F1: organicClosedB3F1,
        closedSingle: organicClosedSingle,
        closedTotal: organicClosedTotal
      },
      todaysClosed: todaysClosed,
      totalSales: totalSales
    }, currentTotals); // Pass the old totals for calculation
    
    return { success: true, message: "Data saved successfully to both Sheet1 and Daily Report! Base totals updated." };
    
  } catch (error) {
    console.error("Error in submitData:", error);
    return { success: false, message: "Error saving data: " + error.toString() };
  }
}

// Modified to use passed totals instead of calculating from Sheet1
function appendToDailyReport(spreadsheet, data, previousTotals) {
  // Get or create the Daily Report sheet
  var reportSheet = spreadsheet.getSheetByName("Daily Report");
  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet("Daily Report");
  }
  
  // Use the passed previous totals (before today's data was added)
  var totals = previousTotals;
  
  // Find the next available row (after previous reports)
  var lastRow = reportSheet.getLastRow();
  var startRow = lastRow + 2; // Leave one empty row between reports
  
  // Create the daily report structure starting from the next available row
  createDailyReportSection(reportSheet, startRow, data, totals);
}

function createDailyReportSection(sheet, startRow, data, totals) {
  // Title and date
  sheet.getRange(startRow, 1, 1, 8).merge();
  sheet.getRange(startRow, 1).setValue("Mandarin Body Wash Daily Enquiry");
  
  sheet.getRange(startRow + 1, 1, 1, 8).merge();
  sheet.getRange(startRow + 1, 1).setValue(data.date);
  
  // Main table headers
  sheet.getRange(startRow + 3, 1).setValue("FB Ad (Messenger)");
  sheet.getRange(startRow + 3, 2).setValue("+/-");
  sheet.getRange(startRow + 3, 3).setValue("Total");
  sheet.getRange(startRow + 3, 5).setValue("Organic");
  sheet.getRange(startRow + 3, 6).setValue("+/-");
  sheet.getRange(startRow + 3, 7).setValue("Total");
  
  // FB Ad section (main table)
  var fbRows = [
    ["Contact:", formatChange(data.fb.enquiry), totals.fb.enquiry + data.fb.enquiry],
    ["Follow Up:", formatChange(data.fb.followup), totals.fb.followup + data.fb.followup],
    ["Waiting Payment:", formatChange(data.fb.waiting), totals.fb.waiting + data.fb.waiting],
    ["No Reply:", formatChange(data.fb.noreply), totals.fb.noreply + data.fb.noreply],
    ["Drop:", formatChange(data.fb.drop), totals.fb.drop + data.fb.drop],
    ["Closed:", formatChange(data.fb.closedCustomers), totals.fb.closedCustomers + data.fb.closedCustomers]
  ];
  
  // Organic section (main table)
  var organicRows = [
    ["Contact:", formatChange(data.organic.enquiry), totals.organic.enquiry + data.organic.enquiry],
    ["Follow Up:", formatChange(data.organic.followup), totals.organic.followup + data.organic.followup],
    ["Waiting Payment:", formatChange(data.organic.waiting), totals.organic.waiting + data.organic.waiting],
    ["No Reply:", formatChange(data.organic.noreply), totals.organic.noreply + data.organic.noreply],
    ["Drop:", formatChange(data.organic.drop), totals.organic.drop + data.organic.drop],
    ["Closed:", formatChange(data.organic.closedCustomers), totals.organic.closedCustomers + data.organic.closedCustomers]
  ];
  
  // Fill FB and Organic data
  for (var i = 0; i < fbRows.length; i++) {
    var row = startRow + 4 + i;
    sheet.getRange(row, 1).setValue(fbRows[i][0]);
    sheet.getRange(row, 2).setValue(fbRows[i][1]);
    sheet.getRange(row, 3).setValue(fbRows[i][2]);
    
    sheet.getRange(row, 5).setValue(organicRows[i][0]);
    sheet.getRange(row, 6).setValue(organicRows[i][1]);
    sheet.getRange(row, 7).setValue(organicRows[i][2]);
  }
  
  // Sales section headers
  var salesStartRow = startRow + 11;
  sheet.getRange(salesStartRow, 1, 1, 4).merge();
  sheet.getRange(salesStartRow, 1).setValue("Total Sales (RM)");
  sheet.getRange(salesStartRow, 5, 1, 4).merge();
  sheet.getRange(salesStartRow, 5).setValue("Total Sales (RM)");
  
  // Calculate sales data
  var fbTotalB3F1 = totals.fb.closedB3F1 + data.fb.closedB3F1;
  var fbTotalSingle = totals.fb.closedSingle + data.fb.closedSingle;
  var fbB3F1Amount = fbTotalB3F1 * 487;
  var fbSingleAmount = fbTotalSingle * 167;
  var fbTotalAmount = fbB3F1Amount + fbSingleAmount;
  
  var organicTotalB3F1 = totals.organic.closedB3F1 + data.organic.closedB3F1;
  var organicTotalSingle = totals.organic.closedSingle + data.organic.closedSingle;
  var organicB3F1Amount = organicTotalB3F1 * 487;
  var organicSingleAmount = organicTotalSingle * 167;
  var organicTotalAmount = organicB3F1Amount + organicSingleAmount;
  
  // FB Sales data
  sheet.getRange(salesStartRow + 1, 1).setValue("Total B3F1 Set Order\n( RM 487 / set )");
  sheet.getRange(salesStartRow + 1, 2).setValue(formatChange(data.fb.closedB3F1));
  sheet.getRange(salesStartRow + 1, 3).setValue(fbTotalB3F1);
  sheet.getRange(salesStartRow + 1, 4).setValue(fbB3F1Amount.toFixed(2));
  
  sheet.getRange(salesStartRow + 2, 1).setValue("Total Single Bottle\n( RM 167 / bottle )");
  sheet.getRange(salesStartRow + 2, 2).setValue(formatChange(data.fb.closedSingle));
  sheet.getRange(salesStartRow + 2, 3).setValue(fbTotalSingle);
  sheet.getRange(salesStartRow + 2, 4).setValue(fbSingleAmount.toFixed(2));
  
  // Calculate percentage using CUSTOMERS not B3F1 sets
  var fbTotalEnquiry = totals.fb.enquiry + data.fb.enquiry;
  var fbTotalClosedCustomers = totals.fb.closedCustomers + data.fb.closedCustomers;
  var fbClosedPercentage = fbTotalEnquiry > 0 ? ((fbTotalClosedCustomers / fbTotalEnquiry) * 100).toFixed(2) : 0;
  
  sheet.getRange(salesStartRow + 3, 1).setValue("Total Closed\n" + fbTotalClosedCustomers + "(" + formatChange(data.fb.closedCustomers) + ")/" + fbTotalEnquiry);
  sheet.getRange(salesStartRow + 3, 2).setValue(fbClosedPercentage + "%");
  sheet.getRange(salesStartRow + 3, 4).setValue(fbTotalAmount.toFixed(2));
  
  // Organic Sales data
  sheet.getRange(salesStartRow + 1, 5).setValue("Total B3F1 Set Order\n( RM 487 / set )");
  sheet.getRange(salesStartRow + 1, 6).setValue(formatChange(data.organic.closedB3F1));
  sheet.getRange(salesStartRow + 1, 7).setValue(organicTotalB3F1);
  sheet.getRange(salesStartRow + 1, 8).setValue(organicB3F1Amount.toFixed(2));
  
  sheet.getRange(salesStartRow + 2, 5).setValue("Total Single Bottle\n( RM 167 / bottle )");
  sheet.getRange(salesStartRow + 2, 6).setValue(formatChange(data.organic.closedSingle));
  sheet.getRange(salesStartRow + 2, 7).setValue(organicTotalSingle);
  sheet.getRange(salesStartRow + 2, 8).setValue(organicSingleAmount.toFixed(2));
  
  // Calculate percentage using CUSTOMERS not B3F1 sets
  var organicTotalEnquiry = totals.organic.enquiry + data.organic.enquiry;
  var organicTotalClosedCustomers = totals.organic.closedCustomers + data.organic.closedCustomers;
  var organicClosedPercentage = organicTotalEnquiry > 0 ? ((organicTotalClosedCustomers / organicTotalEnquiry) * 100).toFixed(2) : 0;
  
  sheet.getRange(salesStartRow + 3, 5).setValue("Total Closed\n" + organicTotalClosedCustomers + "(" + formatChange(data.organic.closedCustomers) + ")/" + organicTotalEnquiry);
  sheet.getRange(salesStartRow + 3, 6).setValue(organicClosedPercentage + "%");
  sheet.getRange(salesStartRow + 3, 8).setValue(organicTotalAmount.toFixed(2));
  
  // Summary table
  var summaryStartRow = salesStartRow + 5;
  var dateShort = data.date.split('/')[0] + "/" + data.date.split('/')[1];
  sheet.getRange(summaryStartRow, 1, 1, 2).merge();
  sheet.getRange(summaryStartRow, 1).setValue("Today (" + dateShort + ") 6:00PM");
  sheet.getRange(summaryStartRow, 3).setValue("Total");
  
  // Summary calculations
  var totalEnquiryChange = data.fb.enquiry + data.organic.enquiry;
  var totalFollowupChange = data.fb.followup + data.organic.followup;
  var totalWaitingChange = data.fb.waiting + data.organic.waiting;
  var totalNoreplyChange = data.fb.noreply + data.organic.noreply;
  var totalDropChange = data.fb.drop + data.organic.drop;
  var totalClosedChange = data.fb.closedCustomers + data.organic.closedCustomers;
  
  var summaryData = [
    ["Enquiry (" + formatChange(totalEnquiryChange) + ")", (totals.fb.enquiry + totals.organic.enquiry) + totalEnquiryChange],
    ["Follow Up (" + formatChange(totalFollowupChange) + ")", (totals.fb.followup + totals.organic.followup) + totalFollowupChange],
    ["Waiting Payment (" + formatChange(totalWaitingChange) + ")", (totals.fb.waiting + totals.organic.waiting) + totalWaitingChange],
    ["No reply (" + formatChange(totalNoreplyChange) + ")", (totals.fb.noreply + totals.organic.noreply) + totalNoreplyChange],
    ["Drop (" + formatChange(totalDropChange) + ")", (totals.fb.drop + totals.organic.drop) + totalDropChange],
    ["Closed (" + formatChange(totalClosedChange) + ")", (totals.fb.closedCustomers + totals.organic.closedCustomers) + totalClosedChange]
  ];
  
  for (var i = 0; i < summaryData.length; i++) {
    sheet.getRange(summaryStartRow + 1 + i, 1).setValue(summaryData[i][0]);
    sheet.getRange(summaryStartRow + 1 + i, 3).setValue(summaryData[i][1]);
  }
  
  // Apply formatting to this section
  formatDailyReportSection(sheet, startRow, summaryStartRow + 6);
}

function formatChange(value) {
  if (value > 0) return "+" + value;
  if (value < 0) return value.toString();
  return "+0";
}

// REMOVED: calculateTotals function - no longer needed since we use Config sheet

function formatDailyReportSection(sheet, startRow, endRow) {
  // Set column widths (only need to do this once for the sheet)
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 100);
  
  // Title formatting
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(startRow + 1, 1).setFontSize(12).setHorizontalAlignment("center");
  
  // Header formatting (matching your HTML colors)
  sheet.getRange(startRow + 3, 1, 1, 3).setBackground("#4a6741").setFontColor("white").setFontWeight("bold");
  sheet.getRange(startRow + 3, 5, 1, 3).setBackground("#bf9000").setFontColor("white").setFontWeight("bold");
  
  // Main table backgrounds (matching HTML)
  sheet.getRange(startRow + 4, 1, 6, 3).setBackground("#d9e5d6"); // FB Ad section
  sheet.getRange(startRow + 4, 5, 6, 3).setBackground("#fff2cc"); // Organic section
  
  // Sales section headers
  var salesStartRow = startRow + 11;
  sheet.getRange(salesStartRow, 1).setBackground("#1c4587").setFontColor("white").setFontWeight("bold");
  sheet.getRange(salesStartRow, 5).setBackground("#1c4587").setFontColor("white").setFontWeight("bold");
  
  // Sales section backgrounds
  sheet.getRange(salesStartRow + 1, 1, 3, 4).setBackground("#cfe2f3");
  sheet.getRange(salesStartRow + 1, 5, 3, 4).setBackground("#cfe2f3");
  
  // Summary section background
  var summaryStartRow = salesStartRow + 5;
  sheet.getRange(summaryStartRow, 1, 7, 3).setBackground("#f4cccc");
  
  // Center align numbers
  sheet.getRange(startRow + 4, 2, endRow - startRow, 6).setHorizontalAlignment("center");
  
  // Right align currency amounts
  sheet.getRange(salesStartRow + 1, 4, 3, 1).setHorizontalAlignment("right");
  sheet.getRange(salesStartRow + 1, 8, 3, 1).setHorizontalAlignment("right");
  
  // Add borders to the entire section
  sheet.getRange(startRow, 1, endRow - startRow, 8).setBorder(true, true, true, true, true, true);
}
