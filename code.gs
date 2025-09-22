function doGet() {
  return HtmlService.createHtmlOutputFromFile('form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMonthlyReportSheet(spreadsheet) {
  var monthName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM");
  var sheetName = monthName + " Report";

  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

function getCurrentTotalsFromDailyReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = getMonthlyReportSheet(spreadsheet);

  // Default values - UPDATED to remove COD
  var totals = {
    fb: { enquiry: 1254, waiting: 8, drop: 7, closedCustomers: 34, closedB3F1: 29, closedSingle: 6 },
    organic: { enquiry: 4, waiting: 0, drop: 0, closedCustomers: 4, closedB3F1: 4, closedSingle: 0 }
  };

  if (!reportSheet) {
    console.log('No report sheet found, using default values');
    return totals;
  }

  var lastRow = reportSheet.getLastRow();
  if (lastRow <= 1) {
    console.log('No data in report sheet, using default values');
    return totals;
  }

  // Read more columns to ensure we capture organic data
  var allData = reportSheet.getRange(1, 1, lastRow, 12).getValues();
  console.log('Reading ' + lastRow + ' rows from report sheet');

  for (var i = 0; i < allData.length; i++) {
    var cellValue = allData[i][0]; // Column A (FB Ad data)
    var organicCellValue = allData[i][4]; // Column E (Organic data)

    if (typeof cellValue === 'string') {
      // FB Ad data processing (Column A, values in Column C)
      if (cellValue.includes('Contact:')) {
        totals.fb.enquiry = Number(allData[i][2]) || totals.fb.enquiry;
        console.log('FB Enquiry found: ' + totals.fb.enquiry);
      }
      if (cellValue.includes('Waiting Payment:')) {
        totals.fb.waiting = Number(allData[i][2]) || totals.fb.waiting;
        console.log('FB Waiting found: ' + totals.fb.waiting);
      }
      if (cellValue.includes('Drop:')) {
        totals.fb.drop = Number(allData[i][2]) || totals.fb.drop;
        console.log('FB Drop found: ' + totals.fb.drop);
      }
      if (cellValue.includes('Closed:') && !cellValue.includes('Total')) {
        totals.fb.closedCustomers = Number(allData[i][2]) || totals.fb.closedCustomers;
        console.log('FB Closed Customers found: ' + totals.fb.closedCustomers);
      }

      // FB Sales data processing
      if (cellValue.includes('Total B3F1 Set Order')) {
        totals.fb.closedB3F1 = Number(allData[i][2]) || totals.fb.closedB3F1;
        console.log('FB B3F1 found: ' + totals.fb.closedB3F1);
      }
      if (cellValue.includes('Total Single Bottle')) {
        totals.fb.closedSingle = Number(allData[i][2]) || totals.fb.closedSingle;
        console.log('FB Single found: ' + totals.fb.closedSingle);
      }
    }

    // Organic data processing (Column E, values in Column G)
    if (typeof organicCellValue === 'string' && organicCellValue.trim() !== '') {
      console.log('Checking organic cell: "' + organicCellValue + '" in row ' + (i + 1));
      
      if (organicCellValue.includes('Contact:')) {
        totals.organic.enquiry = Number(allData[i][6]) || totals.organic.enquiry;
        console.log('Organic Enquiry found: ' + totals.organic.enquiry);
      }
      if (organicCellValue.includes('Waiting Payment:')) {
        totals.organic.waiting = Number(allData[i][6]) || totals.organic.waiting;
        console.log('Organic Waiting found: ' + totals.organic.waiting);
      }
      if (organicCellValue.includes('Drop:')) {
        totals.organic.drop = Number(allData[i][6]) || totals.organic.drop;
        console.log('Organic Drop found: ' + totals.organic.drop);
      }
      if (organicCellValue.includes('Closed:') && !organicCellValue.includes('Total')) {
        totals.organic.closedCustomers = Number(allData[i][6]) || totals.organic.closedCustomers;
        console.log('Organic Closed Customers found: ' + totals.organic.closedCustomers);
      }

      // Organic Sales data processing
      if (organicCellValue.includes('Total B3F1 Set Order')) {
        totals.organic.closedB3F1 = Number(allData[i][6]) || totals.organic.closedB3F1;
        console.log('Organic B3F1 found: ' + totals.organic.closedB3F1);
      }
      if (organicCellValue.includes('Total Single Bottle')) {
        totals.organic.closedSingle = Number(allData[i][6]) || totals.organic.closedSingle;
        console.log('Organic Single found: ' + totals.organic.closedSingle);
      }
    }
  }

  console.log('Final totals:', totals);
  return totals;
}

function submitData(data) {
  try {
    var currentTotals = getCurrentTotalsFromDailyReport();
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Create summary sheet (Sheet1) if missing with updated headers - UPDATED to remove COD
    var summarySheet = spreadsheet.getSheetByName("Sheet1");
    if (!summarySheet) {
      summarySheet = spreadsheet.insertSheet("Sheet1");
      summarySheet.appendRow([
        "Date",
        "FB Enquiry", "FB Waiting Payment", "FB Drop", "FB Closed",
        "Organic Enquiry", "Organic Waiting Payment", "Organic Drop", "Organic Closed",
        "Total Sales", "Today Closed",
        "Today Total Enquiry", "Today Waiting Payment", "Today Total Drop", "Today Total Closed"
      ]);
    }

    var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    // Convert input values (allow negative values for deductions) - UPDATED to remove COD
    var fbEnquiry = parseFloat(data.fbEnquiry) || 0;
    var fbWaiting = parseFloat(data.fbWaiting) || 0;
    var fbDrop = parseFloat(data.fbDrop) || 0;
    var fbClosedB3F1 = parseFloat(data.fbClosedB3F1) || 0;
    var fbClosedSingle = parseFloat(data.fbClosedSingle) || 0;
    var fbClosedCustomers = parseFloat(data.fbClosedCustomers) || 0;

    var organicEnquiry = parseFloat(data.organicEnquiry) || 0;
    var organicWaiting = parseFloat(data.organicWaiting) || 0;
    var organicDrop = parseFloat(data.organicDrop) || 0;
    var organicClosedB3F1 = parseFloat(data.organicClosedB3F1) || 0;
    var organicClosedSingle = parseFloat(data.organicClosedSingle) || 0;
    var organicClosedCustomers = parseFloat(data.organicClosedCustomers) || 0;

    var fbClosedTotal = fbClosedB3F1 + fbClosedSingle;
    var organicClosedTotal = organicClosedB3F1 + organicClosedSingle;

    // Separate today's closed values for clarity
    var fbTodaysClosed = fbClosedCustomers;
    var organicTodaysClosed = organicClosedCustomers;
    var combinedTodaysClosed = fbClosedCustomers + organicClosedCustomers;
    
    var totalSales = (fbClosedB3F1 * 487) + (fbClosedSingle * 167) +
                     (organicClosedB3F1 * 487) + (organicClosedSingle * 167);

    // Calculate today's changes for daily input - UPDATED to remove COD
    var todayTotalEnquiry = fbEnquiry + organicEnquiry;
    var todayWaitingPayment = fbWaiting + organicWaiting;
    var todayTotalDrop = fbDrop + organicDrop;
    var todayTotalClosed = fbClosedCustomers + organicClosedCustomers;

    // Calculate CUMULATIVE totals (current totals + today's changes/deductions) - UPDATED to remove COD
    var cumulativeTotalEnquiry = (currentTotals.fb.enquiry + currentTotals.organic.enquiry) + todayTotalEnquiry;
    var cumulativeWaitingPayment = (currentTotals.fb.waiting + currentTotals.organic.waiting) + todayWaitingPayment;
    var cumulativeTotalDrop = (currentTotals.fb.drop + currentTotals.organic.drop) + todayTotalDrop;
    var cumulativeTotalClosed = (currentTotals.fb.closedCustomers + currentTotals.organic.closedCustomers) + todayTotalClosed;

    // Ensure cumulative values don't go below 0 - UPDATED to remove COD
    cumulativeTotalEnquiry = Math.max(0, cumulativeTotalEnquiry);
    cumulativeWaitingPayment = Math.max(0, cumulativeWaitingPayment);
    cumulativeTotalDrop = Math.max(0, cumulativeTotalDrop);
    cumulativeTotalClosed = Math.max(0, cumulativeTotalClosed);
    
    // UPDATED: Map the data to the correct columns - UPDATED to remove COD
    summarySheet.appendRow([
      date,
      // FB Ad columns - UPDATED to remove COD
      fbEnquiry,              // Column B: Enquiry 
      fbWaiting,              // Column C: Waiting Payment  
      fbDrop,                 // Column D: Drop
      fbClosedCustomers,      // Column E: Closed (FB customers closed today)
      // Organic/WhatsApp columns - UPDATED to remove COD
      organicEnquiry,         // Column F: Enquiry(Whatsapp)
      organicWaiting,         // Column G: No Reply(Whatsapp) 
      organicDrop,            // Column H: Drop(Whatsapp)
      organicClosedCustomers, // Column I: Closed(Whatsapp) (Organic customers closed today)
      // Summary columns - UPDATED column positions due to removed COD
      totalSales,             // Column J: Total Sales
      combinedTodaysClosed,   // Column K: Today Closed
      cumulativeTotalEnquiry, // Column L: Today Total Enquiry
      cumulativeWaitingPayment, // Column M: Today Waiting Payment
      cumulativeTotalDrop,    // Column N: Today Total Drop
      cumulativeTotalClosed   // Column O: Today Total Closed
    ]);

    // Append to monthly report (keeping the ORIGINAL logic) - UPDATED to remove COD
    appendToDailyReport(spreadsheet, {
      date: today,
      fb: { enquiry: fbEnquiry, waiting: fbWaiting, drop: fbDrop, closedCustomers: fbClosedCustomers, closedB3F1: fbClosedB3F1, closedSingle: fbClosedSingle, closedTotal: fbClosedTotal },
      organic: { enquiry: organicEnquiry, waiting: organicWaiting, drop: organicDrop, closedCustomers: organicClosedCustomers, closedB3F1: organicClosedB3F1, closedSingle: organicClosedSingle, closedTotal: organicClosedTotal },
      fbTodaysClosed: fbTodaysClosed,
      organicTodaysClosed: organicTodaysClosed,
      combinedTodaysClosed: combinedTodaysClosed,
      totalSales: totalSales
    }, currentTotals);

    return { success: true, message: "Data saved successfully! FB Closed: " + fbTodaysClosed + ", Organic Closed: " + organicTodaysClosed + ", Total: " + combinedTodaysClosed };

  } catch (error) {
    console.error("Error in submitData:", error);
    return { success: false, message: "Error saving data: " + error.toString() };
  }
}

function appendToDailyReport(spreadsheet, data, previousTotals) {
  var reportSheet = getMonthlyReportSheet(spreadsheet);
  var totals = previousTotals;
  var lastRow = reportSheet.getLastRow();
  var startRow = lastRow + 2;
  createDailyReportSection(reportSheet, startRow, data, totals);
}

function createDailyReportSection(sheet, startRow, data, totals) {
  // Title and date
  sheet.getRange(startRow, 1, 1, 8).merge();
  sheet.getRange(startRow, 1).setValue("Mandarin Body Wash Daily Enquiry").setHorizontalAlignment("center");
  
  sheet.getRange(startRow + 1, 1, 1, 8).merge();
  sheet.getRange(startRow + 1, 1).setValue(data.date).setHorizontalAlignment("center");
  
  // Main table headers (8 columns layout)
  sheet.getRange(startRow + 3, 1).setValue("FB Ad (Messenger)").setHorizontalAlignment("center");
  sheet.getRange(startRow + 3, 2).setValue("+/-").setHorizontalAlignment("center");
  sheet.getRange(startRow + 3, 3).setValue("Total").setHorizontalAlignment("center");
  sheet.getRange(startRow + 3, 4).setValue("--").setHorizontalAlignment("center");
  sheet.getRange(startRow + 3, 5).setValue("Organic").setHorizontalAlignment("center");
  sheet.getRange(startRow + 3, 6).setValue("+/-").setHorizontalAlignment("center");
  sheet.getRange(startRow + 3, 7).setValue("Total").setHorizontalAlignment("center");
  sheet.getRange(startRow + 3, 8).setValue("--").setHorizontalAlignment("center");
  
  // Calculate new totals with safety checks - UPDATED to remove COD
  var newFbEnquiry = Math.max(0, totals.fb.enquiry + data.fb.enquiry);
  var newFbWaiting = Math.max(0, totals.fb.waiting + data.fb.waiting);
  var newFbDrop = Math.max(0, totals.fb.drop + data.fb.drop);
  var newFbClosed = Math.max(0, totals.fb.closedCustomers + data.fb.closedCustomers);
  
  var newOrganicEnquiry = Math.max(0, totals.organic.enquiry + data.organic.enquiry);
  var newOrganicWaiting = Math.max(0, totals.organic.waiting + data.organic.waiting);
  var newOrganicDrop = Math.max(0, totals.organic.drop + data.organic.drop);
  var newOrganicClosed = Math.max(0, totals.organic.closedCustomers + data.organic.closedCustomers);
  
  // FB Ad section (main table) - UPDATED to remove COD
  var fbRows = [
    ["Contact:", formatChange(data.fb.enquiry), newFbEnquiry, "--"],
    ["Waiting Payment:", formatChange(data.fb.waiting), newFbWaiting, "--"],
    ["Drop:", formatChange(data.fb.drop), newFbDrop, "--"],
    ["Closed:", formatChange(data.fb.closedCustomers), newFbClosed, "--"]
  ];
  
  // Organic section (main table) - UPDATED to remove COD
  var organicRows = [
    ["Contact:", formatChange(data.organic.enquiry), newOrganicEnquiry, "--"],
    ["Waiting Payment:", formatChange(data.organic.waiting), newOrganicWaiting, "--"],
    ["Drop:", formatChange(data.organic.drop), newOrganicDrop, "--"],
    ["Closed:", formatChange(data.organic.closedCustomers), newOrganicClosed, "--"]
  ];
  
  // Fill FB and Organic data with center alignment - UPDATED for 4 rows (back to original)
  for (var i = 0; i < fbRows.length; i++) {
    var row = startRow + 4 + i;
    sheet.getRange(row, 1).setValue(fbRows[i][0]).setHorizontalAlignment("center");
    sheet.getRange(row, 2).setValue(fbRows[i][1]).setHorizontalAlignment("center");
    sheet.getRange(row, 3).setValue(fbRows[i][2]).setHorizontalAlignment("center");
    sheet.getRange(row, 4).setValue(fbRows[i][3]).setHorizontalAlignment("center");
    
    sheet.getRange(row, 5).setValue(organicRows[i][0]).setHorizontalAlignment("center");
    sheet.getRange(row, 6).setValue(organicRows[i][1]).setHorizontalAlignment("center");
    sheet.getRange(row, 7).setValue(organicRows[i][2]).setHorizontalAlignment("center");
    sheet.getRange(row, 8).setValue(organicRows[i][3]).setHorizontalAlignment("center");
  }
  
  // Sales section headers - UPDATED row position back to original
  var salesStartRow = startRow + 9;  // Changed back from 10 to 9
  sheet.getRange(salesStartRow, 1, 1, 4).merge();
  sheet.getRange(salesStartRow, 1).setValue("Total Sales (RM)").setHorizontalAlignment("center");
  sheet.getRange(salesStartRow, 5, 1, 4).merge();
  sheet.getRange(salesStartRow, 5).setValue("Total Sales (RM)").setHorizontalAlignment("center");
  
  // Calculate sales data with safety checks
  var fbTotalB3F1 = Math.max(0, totals.fb.closedB3F1 + data.fb.closedB3F1);
  var fbTotalSingle = Math.max(0, totals.fb.closedSingle + data.fb.closedSingle);
  var fbB3F1Amount = fbTotalB3F1 * 487;
  var fbSingleAmount = fbTotalSingle * 167;
  var fbTotalAmount = fbB3F1Amount + fbSingleAmount;
  
  var organicTotalB3F1 = Math.max(0, totals.organic.closedB3F1 + data.organic.closedB3F1);
  var organicTotalSingle = Math.max(0, totals.organic.closedSingle + data.organic.closedSingle);
  var organicB3F1Amount = organicTotalB3F1 * 487;
  var organicSingleAmount = organicTotalSingle * 167;
  var organicTotalAmount = organicB3F1Amount + organicSingleAmount;
  
  // FB Sales data
  sheet.getRange(salesStartRow + 1, 1).setValue("Total B3F1 Set Order\n( RM 487 / set )").setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 1, 2).setValue(formatChange(data.fb.closedB3F1)).setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 1, 3).setValue(fbTotalB3F1).setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 1, 4).setValue(fbB3F1Amount.toFixed(2)).setHorizontalAlignment("center");
  
  sheet.getRange(salesStartRow + 2, 1).setValue("Total Single Bottle\n( RM 167 / bottle )").setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 2, 2).setValue(formatChange(data.fb.closedSingle)).setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 2, 3).setValue(fbTotalSingle).setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 2, 4).setValue(fbSingleAmount.toFixed(2)).setHorizontalAlignment("center");
  
  // Calculate percentage using CUSTOMERS not B3F1 sets
  var fbClosedPercentage = newFbEnquiry > 0 ? ((newFbClosed / newFbEnquiry) * 100).toFixed(2) : 0;
  
  sheet.getRange(salesStartRow + 3, 1).setValue("Total Closed\n" + newFbClosed + "(" + formatChange(data.fb.closedCustomers) + ")/" + newFbEnquiry).setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 3, 2).setValue(fbClosedPercentage + "%").setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 3, 4).setValue(fbTotalAmount.toFixed(2)).setHorizontalAlignment("center");
  
  // Organic Sales data
  sheet.getRange(salesStartRow + 1, 5).setValue("Total B3F1 Set Order\n( RM 487 / set )").setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 1, 6).setValue(formatChange(data.organic.closedB3F1)).setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 1, 7).setValue(organicTotalB3F1).setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 1, 8).setValue(organicB3F1Amount.toFixed(2)).setHorizontalAlignment("center");
  
  sheet.getRange(salesStartRow + 2, 5).setValue("Total Single Bottle\n( RM 167 / bottle )").setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 2, 6).setValue(formatChange(data.organic.closedSingle)).setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 2, 7).setValue(organicTotalSingle).setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 2, 8).setValue(organicSingleAmount.toFixed(2)).setHorizontalAlignment("center");
  
  // Calculate percentage using CUSTOMERS not B3F1 sets
  var organicClosedPercentage = newOrganicEnquiry > 0 ? ((newOrganicClosed / newOrganicEnquiry) * 100).toFixed(2) : 0;
  
  sheet.getRange(salesStartRow + 3, 5).setValue("Total Closed\n" + newOrganicClosed + "(" + formatChange(data.organic.closedCustomers) + ")/" + newOrganicEnquiry).setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 3, 6).setValue(organicClosedPercentage + "%").setHorizontalAlignment("center");
  sheet.getRange(salesStartRow + 3, 8).setValue(organicTotalAmount.toFixed(2)).setHorizontalAlignment("center");
  
  // *** TODAY SUMMARY FORMAT - UPDATED to remove COD ***
  var summaryStartRow = startRow + 3;
  var summaryColumn = 10; // Column J
  
  var dateShort = data.date.split('/')[0] + "/" + data.date.split('/')[1];
  
  // Original format: "Today (dd/mm) 6:00PM" header merged across 2 columns, then "Total" column
  sheet.getRange(summaryStartRow, summaryColumn, 1, 2).merge();
  sheet.getRange(summaryStartRow, summaryColumn).setValue("Today (" + dateShort + ") 6:00PM").setHorizontalAlignment("center");
  sheet.getRange(summaryStartRow, summaryColumn + 2).setValue("Total").setHorizontalAlignment("center");
  
  // Calculate summary totals (cumulative totals after today's changes) - UPDATED to remove COD
  var summaryTotalEnquiry = newFbEnquiry + newOrganicEnquiry;
  var summaryTotalWaiting = newFbWaiting + newOrganicWaiting;
  var summaryTotalDrop = newFbDrop + newOrganicDrop;
  var summaryTotalClosed = newFbClosed + newOrganicClosed;
  
  // Original summary data format (3 columns: Label, Change, Total) - UPDATED to remove COD
  var summaryRows = [
    ["Enquiry (" + formatChange(data.fb.enquiry + data.organic.enquiry) + ")", "", summaryTotalEnquiry],
    ["Waiting Payment (" + formatChange(data.fb.waiting + data.organic.waiting) + ")", "", summaryTotalWaiting],
    ["Drop (" + formatChange(data.fb.drop + data.organic.drop) + ")", "", summaryTotalDrop],
    ["Closed (" + formatChange(data.fb.closedCustomers + data.organic.closedCustomers) + ")", "", summaryTotalClosed]
  ];
  
  // Fill original summary data format - UPDATED for 4 rows (back to original)
  for (var i = 0; i < summaryRows.length; i++) {
    var row = summaryStartRow + 1 + i;
    sheet.getRange(row, summaryColumn, 1, 2).merge(); // Merge first two columns for the label
    sheet.getRange(row, summaryColumn).setValue(summaryRows[i][0]).setHorizontalAlignment("center");
    sheet.getRange(row, summaryColumn + 2).setValue(summaryRows[i][2]).setHorizontalAlignment("center");
  }

  // Apply formatting - UPDATED row count back to original
  formatDailyReportSection(sheet, startRow, summaryStartRow + 4);  // Changed back from 5 to 4
}

function formatChange(value) {
  // Handle decimal values and ensure proper formatting
  if (value > 0) return "+" + value;
  if (value < 0) return value.toString(); // Will already have negative sign
  return "+0";
}

function formatDailyReportSection(sheet, startRow, endRow) {
  // Set column widths
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 80);
  sheet.setColumnWidth(10, 150);
  sheet.setColumnWidth(11, 80);
  sheet.setColumnWidth(12, 80);

  // Title formatting
  sheet.getRange(startRow, 1).setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(startRow + 1, 1).setFontSize(12).setHorizontalAlignment("center");
  
  // Header formatting
  sheet.getRange(startRow + 3, 1, 1, 4).setBackground("#346855").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(startRow + 3, 5, 1, 4).setBackground("#346855").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Original summary header formatting (3 columns: merged header + Total column)
  var summaryStartRow = startRow + 3;
  sheet.getRange(summaryStartRow, 10, 1, 2).setBackground("#f4cccc").setFontColor("black").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(summaryStartRow, 12).setBackground("#f4cccc").setFontColor("black").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Main table backgrounds - UPDATED for 4 rows (back to original)
  sheet.getRange(startRow + 4, 1, 4, 4).setBackground("#b7e1cd"); // FB Ad section
  sheet.getRange(startRow + 4, 5, 4, 4).setBackground("#fce8b2"); // Organic section
  
  // Sales section headers - UPDATED row position back to original
  var salesStartRow = startRow + 9;  // Changed back from 10 to 9
  sheet.getRange(salesStartRow, 1).setBackground("#c9daf8").setFontColor("black").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(salesStartRow, 5).setBackground("#d9d2e9").setFontColor("black").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Sales section backgrounds
  sheet.getRange(salesStartRow + 1, 1, 3, 4).setBackground("#c9daf8");
  sheet.getRange(salesStartRow + 1, 5, 3, 4).setBackground("#d9d2e9");
  
  // Add borders
  sheet.getRange(salesStartRow, 1, 4, 4).setBorder(true, true, true, true, true, true); // FB Sales section
  sheet.getRange(salesStartRow, 5, 4, 4).setBorder(true, true, true, true, true, true); // Organic Sales section
  sheet.getRange(summaryStartRow, 10, 5, 3).setBorder(true, true, true, true, true, true); // Original summary section (3 columns, 5 rows now back to 4)
  
  // Center align all content
  sheet.getRange(startRow, 1, endRow - startRow + 10, 16).setHorizontalAlignment("center");
}
