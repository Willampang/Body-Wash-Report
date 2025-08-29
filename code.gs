function doGet() {
  return HtmlService.createHtmlOutputFromFile('form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Calculate totals from Daily Report sheet data
function getCurrentTotalsFromDailyReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = spreadsheet.getSheetByName("Daily Report");
  
  // Default base values
  var totals = {
    fb: { enquiry: 1254, followup: 688, waiting: 8, noreply: 536, drop: 7, closedCustomers: 34, closedB3F1: 29, closedSingle: 6 },
    organic: { enquiry: 5, followup: 0, waiting: 0, noreply: 0, drop: 0, closedCustomers: 5, closedB3F1: 5, closedSingle: 0 }
  };
  
  if (!reportSheet) {
    return totals; // Return defaults if no Daily Report sheet exists
  }
  
  var lastRow = reportSheet.getLastRow();
  if (lastRow <= 1) {
    return totals; // Return defaults if no data
  }
  
  // Get all data from the Daily Report sheet
  var allData = reportSheet.getRange(1, 1, lastRow, 20).getValues();
  
  // Parse through the Daily Report data to extract accumulated totals
  for (var i = 0; i < allData.length; i++) {
    var cellValue = allData[i][0];
    
    // Look for rows that contain total data (after each daily report)
    if (typeof cellValue === 'string') {
      // Check for FB Ad data rows
      if (cellValue.includes('Contact:')) {
        var totalCol = allData[i][2];
        if (typeof totalCol === 'number') {
          totals.fb.enquiry = totalCol;
        }
      } else if (cellValue.includes('Follow Up:')) {
        var totalCol = allData[i][2];
        if (typeof totalCol === 'number') {
          totals.fb.followup = totalCol;
        }
      } else if (cellValue.includes('Waiting Payment:')) {
        var totalCol = allData[i][2];
        if (typeof totalCol === 'number') {
          totals.fb.waiting = totalCol;
        }
      } else if (cellValue.includes('No Reply:')) {
        var totalCol = allData[i][2];
        if (typeof totalCol === 'number') {
          totals.fb.noreply = totalCol;
        }
      } else if (cellValue.includes('Drop:')) {
        var totalCol = allData[i][2];
        if (typeof totalCol === 'number') {
          totals.fb.drop = totalCol;
        }
      } else if (cellValue.includes('Closed:') && !cellValue.includes('Total')) {
        var totalCol = allData[i][2];
        if (typeof totalCol === 'number') {
          totals.fb.closedCustomers = totalCol;
        }
      }
      
      // Check for Organic data (column 6)
      if (i < allData.length && allData[i][4]) {
        var organicCell = allData[i][4];
        if (typeof organicCell === 'string') {
          if (organicCell.includes('Contact:')) {
            var organicTotal = allData[i][6];
            if (typeof organicTotal === 'number') {
              totals.organic.enquiry = organicTotal;
            }
          } else if (organicCell.includes('Follow Up:')) {
            var organicTotal = allData[i][6];
            if (typeof organicTotal === 'number') {
              totals.organic.followup = organicTotal;
            }
          } else if (organicCell.includes('Waiting Payment:')) {
            var organicTotal = allData[i][6];
            if (typeof organicTotal === 'number') {
              totals.organic.waiting = organicTotal;
            }
          } else if (organicCell.includes('No Reply:')) {
            var organicTotal = allData[i][6];
            if (typeof organicTotal === 'number') {
              totals.organic.noreply = organicTotal;
            }
          } else if (organicCell.includes('Drop:')) {
            var organicTotal = allData[i][6];
            if (typeof organicTotal === 'number') {
              totals.organic.drop = organicTotal;
            }
          } else if (organicCell.includes('Closed:') && !organicCell.includes('Total')) {
            var organicTotal = allData[i][6];
            if (typeof organicTotal === 'number') {
              totals.organic.closedCustomers = organicTotal;
            }
          }
        }
      }
      
      // Look for B3F1 and Single bottle totals in sales section
      if (cellValue.includes('Total B3F1 Set Order')) {
        var b3f1Total = allData[i][2];
        if (typeof b3f1Total === 'number') {
          totals.fb.closedB3F1 = b3f1Total;
        }
      } else if (cellValue.includes('Total Single Bottle')) {
        var singleTotal = allData[i][2];
        if (typeof singleTotal === 'number') {
          totals.fb.closedSingle = singleTotal;
        }
      }
      
      // Check for Organic sales data (column 4)
      if (i < allData.length && allData[i][4]) {
        var organicSalesCell = allData[i][4];
        if (typeof organicSalesCell === 'string') {
          if (organicSalesCell.includes('Total B3F1 Set Order')) {
            var organicB3F1Total = allData[i][6];
            if (typeof organicB3F1Total === 'number') {
              totals.organic.closedB3F1 = organicB3F1Total;
            }
          } else if (organicSalesCell.includes('Total Single Bottle')) {
            var organicSingleTotal = allData[i][6];
            if (typeof organicSingleTotal === 'number') {
              totals.organic.closedSingle = organicSingleTotal;
            }
          }
        }
      }
    }
  }
  
  return totals;
}

function submitData(data) {
  try {
    // Get current totals from Daily Report sheet
    var currentTotals = getCurrentTotalsFromDailyReport();
    
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
    }, currentTotals); // Pass the current totals for calculation
    
    return { success: true, message: "Data saved successfully to both Sheet1 and Daily Report!" };
    
  } catch (error) {
    console.error("Error in submitData:", error);
    return { success: false, message: "Error saving data: " + error.toString() };
  }
}

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
