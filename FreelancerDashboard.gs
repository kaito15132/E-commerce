/**
 * Freelancer & Vendor Project Dashboard
 * Seller-side Google Apps Script builder.
 *
 * Run buildFreelancerVendorDashboard() from the master Google Sheet.
 * The generated buyer-facing spreadsheet uses only formulas, formatting,
 * dropdowns, conditional formatting, and charts. No buyer scripts required.
 */

const FVD = {
  sheets: [
    'Start Here',
    'Dashboard',
    'Calendar',
    'Projects',
    'Vendors',
    'Deliverables',
    'Payments',
    'Contracts',
    'Assets',
    'Communication Log',
    'Revisions',
    'Setup'
  ],
  colors: {
    darkNavy: '#1F2D3A',
    nearBlackNavy: '#151B2D',
    accentBlue: '#2F80ED',
    softTeal: '#57C7B6',
    softCoral: '#F28B82',
    softYellow: '#FFF3CD',
    softGreen: '#DDEFE3',
    softRed: '#F8D7DA',
    softBeige: '#EED6CC',
    lightGray: '#F6F8FA',
    white: '#FFFFFF',
    textDark: '#202A33',
    mutedText: '#6B7280',
    borderGray: '#D9DEE3'
  },
  calendarColors: {
    project: '#DDEFE3',
    deliverable: '#DCEBFF',
    payment: '#F28B82',
    contract: '#EED6CC',
    revision: '#57C7B6',
    dateBand: '#F6F8FA'
  },
  font: 'Arial',
  firstDataRow: 6,
  maxRows: 300
};

function buildFreelancerVendorDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  resetWorkbook(ss);

  createSetupSheet(ss);
  createStartHereSheet(ss);
  createProjectsSheet(ss);
  createVendorsSheet(ss);
  createDeliverablesSheet(ss);
  createPaymentsSheet(ss);
  createContractsSheet(ss);
  createAssetsSheet(ss);
  createCommunicationLogSheet(ss);
  createRevisionsSheet(ss);
  createDashboardSheet(ss);
  createCalendarSheet(ss);

  addSampleData(ss);
  addDropdowns(ss);
  addFormulas(ss);
  addConditionalFormatting(ss);
  SpreadsheetApp.flush();
  addDashboardCharts(ss);
  finalizeWorkbook(ss);
}

function resetWorkbook(ss) {
  const placeholder = ss.insertSheet(`TEMP_BUILDER_DELETE_${Date.now()}`);
  ss.getSheets().forEach(sheet => {
    if (sheet.getName() !== placeholder.getName()) ss.deleteSheet(sheet);
  });
  FVD.sheets.forEach(name => ss.insertSheet(name));
  ss.deleteSheet(placeholder);
}

function createStartHereSheet(ss) {
  const sheet = ss.getSheetByName('Start Here');
  formatSheet(sheet, 40, 8);
  sheet.setTabColor(FVD.colors.accentBlue);
  sheet.setColumnWidths(1, 1, 32);
  sheet.setColumnWidths(2, 6, 150);
  sheet.setRowHeights(1, 2, 42);

  sheet.getRange('B2:G3').merge()
    .setValue('Freelancer & Vendor Project Dashboard')
    .setFontSize(26).setFontWeight('bold').setFontColor(FVD.colors.white)
    .setBackground(FVD.colors.darkNavy).setVerticalAlignment('middle');
  sheet.getRange('B4:G4').merge()
    .setValue('A premium Google Sheets command center for managing outsourced work, vendors, deadlines, contracts, payments, assets, communication, and revisions.')
    .setFontSize(11).setFontColor(FVD.colors.mutedText).setBackground(FVD.colors.white);

  const sections = [
    ['Welcome', 'Thank you for purchasing this digital Google Sheets template. Start by reviewing the sample data, then replace it with your own projects, vendors, payments, deliverables, contracts, and deadlines.'],
    ['What This Template Helps You Manage', 'Projects, freelancers, vendors, deliverables, payments, contracts, shared assets, communication history, revisions, budget health, and upcoming deadlines.'],
    ['How To Use The Template', 'Edit the tracking tabs first: Projects, Vendors, Deliverables, Payments, Contracts, Assets, Communication Log, and Revisions. The Dashboard and Calendar update from those tabs.'],
    ['Tabs To Edit', 'Projects, Vendors, Deliverables, Payments, Contracts, Assets, Communication Log, Revisions, and Setup. Replace the sample rows with your own records.'],
    ['Dashboard & View Tabs', 'Dashboard and Calendar are designed as summary views. You can edit them, but they include formulas and helper areas that power the visual dashboard.'],
    ['Customize Dropdowns', 'Go to the Setup tab to edit dropdown values. Dropdowns in the tracker tabs pull from the Setup tab lists.'],
    ['Digital Template Reminder', 'This is a digital Google Sheets template. No physical item is shipped. Make your own copy before editing if you want to preserve the original.'],
    ['No Scripts Required', 'The buyer-facing version does not require scripts, add-ons, custom menus, or authorization. Everything works with built-in Google Sheets features.']
  ];

  let row = 7;
  sections.forEach(section => {
    sheet.getRange(row, 2, 1, 6).merge()
      .setValue(section[0])
      .setFontSize(14).setFontWeight('bold').setFontColor(FVD.colors.darkNavy)
      .setBackground(FVD.colors.softBeige)
      .setBorder(true, true, true, true, null, null, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(row + 1, 2, 2, 6).merge()
      .setValue(section[1])
      .setWrap(true).setVerticalAlignment('top').setFontSize(11)
      .setFontColor(FVD.colors.textDark).setBackground(FVD.colors.white)
      .setBorder(true, true, true, true, null, null, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
    row += 4;
  });
}

function createDashboardSheet(ss) {
  const sheet = ss.getSheetByName('Dashboard');
  formatSheet(sheet, 110, 44);
  sheet.setTabColor(FVD.colors.darkNavy);
  sheet.setFrozenRows(5);
  sheet.setColumnWidths(1, 1, 24);
  sheet.setColumnWidths(2, 16, 92);
  sheet.setColumnWidths(28, 10, 125);
  sheet.setRowHeights(8, 11, 30);
  sheet.setRowHeights(22, 27, 28);
  sheet.setRowHeights(53, 14, 36);

  sheet.getRange('B2:Q4').merge()
    .setValue('Freelancer & Vendor Project Dashboard')
    .setFontSize(26).setFontWeight('bold').setFontColor(FVD.colors.white)
    .setBackground(FVD.colors.nearBlackNavy).setVerticalAlignment('middle');
  sheet.getRange('B5:Q5').merge()
    .setValue('Track outsourced projects, vendors, deliverables, payments, contracts, revisions, and deadlines in one simple command center.')
    .setFontSize(11).setFontColor(FVD.colors.white).setBackground(FVD.colors.nearBlackNavy);
  sheet.getRange('B6:Q6').merge()
    .setValue('Template preview uses sample data. Replace with your own information.')
    .setFontSize(10).setFontColor(FVD.colors.mutedText).setBackground(FVD.colors.lightGray);

  const cards = [
    ['Active Projects', 'Projects currently moving', FVD.colors.accentBlue, '#EAF3FF'],
    ['Overdue Deliverables', 'Needs immediate follow-up', FVD.colors.softCoral, FVD.colors.softRed],
    ['Payments Due', 'Due today or earlier', FVD.colors.softTeal, '#EAF8F6'],
    ['Unpaid Balance', 'Open scheduled, due, or overdue', FVD.colors.softCoral, '#FFF1EF'],
    ['Monthly Vendor Spend', 'Paid or due this month', FVD.colors.softTeal, '#EAF8F6'],
    ['Projects Over Budget', 'Actual spend above budget', FVD.colors.softCoral, FVD.colors.softRed],
    ['Missing Contracts', 'Missing, needed, or not started', FVD.colors.softBeige, '#F7ECE7'],
    ['Average Vendor Rating', 'Average of rated vendors', FVD.colors.accentBlue, '#EAF3FF']
  ];
  const cardPositions = ['B8:E12', 'F8:I12', 'J8:M12', 'N8:Q12', 'B14:E18', 'F14:I18', 'J14:M18', 'N14:Q18'];
  cardPositions.forEach((a1, i) => styleKpiCard(sheet, a1, cards[i]));

  sectionTitle(sheet, 'B21:Q21', 'Dashboard Snapshot');
  stylePanel(sheet, 'B22:I34');
  stylePanel(sheet, 'J22:Q34');
  stylePanel(sheet, 'B36:I48');
  stylePanel(sheet, 'J36:Q48');

  sectionTitle(sheet, 'B51:H51', 'Priority Watchlist');
  styleTableHeader(sheet.getRange('B52:H52'));
  sheet.getRange('B52:H52').setValues([['Priority', 'Item', 'Type', 'Vendor', 'Due Date', 'Status', 'Action Needed']]);
  sheet.setRowHeight(51, 26);
  sheet.setRowHeight(52, 34);
  sheet.getRange('B53:H64').setBackground(FVD.colors.white)
    .setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setFontSize(9);
  sheet.getRange('B53:B64').setFontWeight('bold');
  sheet.getRange('F53:F64').setNumberFormat('mmm d');
  sheet.getRange('B51:H64').setBorder(true, true, true, true, null, null, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sectionTitle(sheet, 'J51:O51', 'Upcoming Deadlines');
  styleTableHeader(sheet.getRange('J52:O52'));
  sheet.getRange('J52:O52').setValues([['Date', 'Type', 'Project', 'Vendor', 'Item', 'Status']]);
  sheet.setRowHeight(51, 26);
  sheet.setRowHeight(52, 34);
  sheet.getRange('J53:O64').setBackground(FVD.colors.white)
    .setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setFontSize(9);
  sheet.getRange('J53:J64').setNumberFormat('mmm d').setFontWeight('bold');
  sheet.getRange('J51:O64').setBorder(true, true, true, true, null, null, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sectionTitle(sheet, 'B68:D68', 'Project Status Snapshot');
  sectionTitle(sheet, 'G68:I68', 'Payment Status Snapshot');
  sectionTitle(sheet, 'L68:N68', 'Deliverable Status Snapshot');
  styleTableHeader(sheet.getRange('B69:D69'));
  styleTableHeader(sheet.getRange('G69:I69'));
  styleTableHeader(sheet.getRange('L69:N69'));
  sheet.getRange('B69:D69').setValues([['Status', 'Count', '']]);
  sheet.getRange('G69:I69').setValues([['Status', 'Count', '']]);
  sheet.getRange('L69:N69').setValues([['Status', 'Count', '']]);
  sheet.getRange('C69:D69').merge();
  sheet.getRange('H69:I69').merge();
  sheet.getRange('M69:N69').merge();
  sheet.getRange('B70:D81').setBackground(FVD.colors.white).setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('G70:I81').setBackground(FVD.colors.white).setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('L70:N81').setBackground(FVD.colors.white).setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('C70:D81').mergeAcross();
  sheet.getRange('H70:I81').mergeAcross();
  sheet.getRange('M70:N81').mergeAcross();

  sheet.getRange('AB1:AK1').merge().setValue('Dashboard Helper Area').setFontWeight('bold').setFontColor(FVD.colors.white).setBackground(FVD.colors.nearBlackNavy);
  sheet.getRange('AB:AK').setBackground(FVD.colors.lightGray);
  sheet.hideColumns(28, 10);
}

function styleKpiCard(sheet, a1, card) {
  const range = sheet.getRange(a1);
  range.setBackground(card[3])
    .setBorder(true, true, true, true, null, null, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  const row = range.getRow();
  const col = range.getColumn();
  sheet.getRange(row, col, 1, 4).merge().setValue(card[0])
    .setFontSize(9).setFontWeight('bold').setFontColor(FVD.colors.white)
    .setBackground(card[2])
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange(row + 1, col, 2, 4).merge()
    .setFontSize(24).setFontWeight('bold').setFontColor(FVD.colors.textDark)
    .setBackground(card[3])
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange(row + 3, col, 2, 4).merge().setValue(card[1])
    .setFontSize(9).setFontColor(FVD.colors.mutedText).setBackground(card[3])
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top')
    .setWrap(true);
}

function createCalendarSheet(ss) {
  const sheet = ss.getSheetByName('Calendar');
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  formatSheet(sheet, 300, 32);
  sheet.setTabColor(FVD.colors.softTeal);
  sheet.setColumnWidths(1, 1, 24);
  sheet.setColumnWidths(2, 7, 145);
  sheet.setColumnWidth(10, 76);
  sheet.setColumnWidth(11, 92);
  sheet.setColumnWidth(12, 190);
  sheet.setColumnWidth(13, 120);
  sheet.setColumnWidths(27, 6, 120);
  sheet.setFrozenRows(8);

  sheet.getRange('B2:H3').merge().setFormula('=$C$6&" "&$F$6&" Deadline Calendar"')
    .setFontSize(24).setFontWeight('bold').setFontColor(FVD.colors.white).setBackground(FVD.colors.darkNavy)
    .setVerticalAlignment('middle');
  sheet.getRange('B4:H4').merge().setValue('Deadlines populate from the Projects, Deliverables, Payments, Contracts, and Revisions tabs.')
    .setFontColor(FVD.colors.mutedText).setBackground(FVD.colors.white);

  sheet.getRange('B6').setValue('Month').setFontWeight('bold').setFontColor(FVD.colors.darkNavy);
  sheet.getRange('C6').setValue(monthNames[new Date().getMonth()]).setBackground(FVD.colors.white);
  sheet.getRange('E6').setValue('Year').setFontWeight('bold').setFontColor(FVD.colors.darkNavy);
  sheet.getRange('F6').setValue(new Date().getFullYear()).setBackground(FVD.colors.white);
  sheet.getRange('B6:F6').setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('C7').setFormula('=MONTH(DATEVALUE($C$6&" 1"))');
  sheet.hideRows(7);

  sheet.getRange('B8:H8').setValues([['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']]);
  styleTableHeader(sheet.getRange('B8:H8'));
  sheet.getRange('B9:H14').setBackground(FVD.colors.white)
    .setFontSize(9)
    .setFontWeight('bold')
    .setWrap(true)
    .setVerticalAlignment('top')
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeights(9, 6, 118);

  // Keep the right sidebar focused on simple month-level context; action tables sit below the calendar.
  sectionTitle(sheet, 'J2:M2', 'Calendar Legend');
  const legend = [
    ['[P]', 'Project deadline or launch date'],
    ['[D]', 'Deliverable due'],
    ['[$]', 'Payment due'],
    ['[C]', 'Contract, NDA, W-9, proposal, or scope deadline'],
    ['[R]', 'Revision due or approval deadline']
  ];
  sheet.getRange(3, 10, legend.length, 1).setValues(legend.map(item => [item[0]]))
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontColor(FVD.colors.darkNavy);
  legend.forEach((item, i) => {
    sheet.getRange(3 + i, 11, 1, 3).merge().setValue(item[1]);
  });
  sheet.getRange('J3:M7')
    .setBackground(FVD.colors.white)
    .setFontColor(FVD.colors.textDark)
    .setFontSize(9)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeights(3, 5, 28);

  sectionTitle(sheet, 'J8:M8', 'Monthly Deadline Summary');
  const summaryLabels = [
    'Total Deadlines',
    'Project Deadlines',
    'Deliverables Due',
    'Payments Due',
    'Contracts Due',
    'Revisions Due'
  ];
  summaryLabels.forEach((label, i) => {
    const row = 9 + i;
    sheet.getRange(row, 10).setFontSize(20).setFontWeight('bold').setFontColor(FVD.colors.darkNavy)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(row, 11, 1, 3).merge().setValue(label)
      .setFontSize(10).setFontColor(FVD.colors.mutedText).setWrap(true)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  });
  sheet.getRange('J9:M14')
    .setBackground(FVD.colors.white)
    .setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);

  sectionTitle(sheet, 'B16:E16', 'Upcoming This Week');
  sheet.getRange('B17:E17').setValues([['Date', 'Type', 'Item', 'Vendor']]);
  styleTableHeader(sheet.getRange('B17:E17'));
  sheet.getRange('B18:E27')
    .setBackground(FVD.colors.white)
    .setFontSize(9)
    .setWrap(false)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('D18:D27').setWrap(true);
  sheet.getRange('B18:B27').setNumberFormat('mmm d').setHorizontalAlignment('center');
  sheet.getRange('C18:C27').setHorizontalAlignment('center');

  // Column F is intentionally left blank to separate the two action tables.
  sectionTitle(sheet, 'G16:J16', 'Overdue Items');
  sheet.getRange('G17:J17').setValues([['Date', 'Type', 'Item', 'Vendor']]);
  styleTableHeader(sheet.getRange('G17:J17'));
  sheet.getRange('G18:J27')
    .setBackground(FVD.colors.white)
    .setFontSize(9)
    .setWrap(false)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('I18:I27').setWrap(true);
  sheet.getRange('G18:G27').setNumberFormat('mmm d').setHorizontalAlignment('center');
  sheet.getRange('H18:H27').setHorizontalAlignment('center');
  sheet.setRowHeight(17, 26);
  sheet.setRowHeights(18, 10, 32);

  sheet.getRange('AA1:AF1').merge().setValue('Calendar Event Helper').setFontWeight('bold').setFontColor(FVD.colors.white).setBackground(FVD.colors.nearBlackNavy);
  sheet.getRange('AA2:AF2').setValues([['Date', 'Type', 'Item', 'Vendor', 'Calendar Label', 'Status']]);
  styleTableHeader(sheet.getRange('AA2:AF2'));
  sheet.getRange('AA3:AF300').setBackground(FVD.colors.lightGray);
  sheet.hideColumns(27, 6);
}

function createProjectsSheet(ss) {
  const headers = ['Project ID', 'Project Name', 'Project Type', 'Client/Business Area', 'Primary Vendor', 'Priority', 'Status', 'Start Date', 'Due/Launch Date', 'Budget', 'Actual Spend', 'Budget Status', 'Progress %', 'Notes'];
  createTrackerSheet(ss.getSheetByName('Projects'), 'Projects', 'Plan and monitor outsourced projects, timelines, budgets, and progress.', headers);
}

function createVendorsSheet(ss) {
  const headers = ['Vendor ID', 'Vendor Name', 'Role', 'Email', 'Phone', 'Website/Portfolio', 'Rate Type', 'Rate', 'Payment Terms', 'Contract Status', 'Average Rating', 'Rehire?', 'Notes'];
  createTrackerSheet(ss.getSheetByName('Vendors'), 'Vendors', 'Store freelancer and vendor details, rates, contracts, ratings, and rehire notes.', headers);
}

function createDeliverablesSheet(ss) {
  const headers = ['Deliverable ID', 'Project ID', 'Project Name', 'Vendor', 'Deliverable', 'Due Date', 'Status', 'Priority', 'Revision Needed?', 'Approval Date', 'Notes'];
  createTrackerSheet(ss.getSheetByName('Deliverables'), 'Deliverables', 'Track each deliverable, due date, status, priority, approvals, and revision needs.', headers);
}

function createPaymentsSheet(ss) {
  const headers = ['Payment ID', 'Project ID', 'Project Name', 'Vendor', 'Payment Type', 'Amount', 'Due Date', 'Paid Date', 'Status', 'Payment Method', 'Invoice Link', 'Notes'];
  createTrackerSheet(ss.getSheetByName('Payments'), 'Payments', 'Monitor vendor payments, due dates, paid dates, methods, invoices, and open balances.', headers);
}

function createContractsSheet(ss) {
  const headers = ['Contract ID', 'Project ID', 'Project Name', 'Vendor', 'Document Type', 'Status', 'Date Requested', 'Due Date', 'Date Received/Signed', 'File Link', 'Notes'];
  createTrackerSheet(ss.getSheetByName('Contracts'), 'Contracts', 'Track contracts, NDAs, W-9s, proposals, scopes, signatures, and missing documents.', headers);
}

function createAssetsSheet(ss) {
  const headers = ['Asset ID', 'Project ID', 'Project Name', 'Vendor', 'Asset Type', 'Asset/Link Name', 'URL or Location', 'Access Level', 'Shared Date', 'Return/Revoke Access?', 'Notes'];
  createTrackerSheet(ss.getSheetByName('Assets'), 'Assets', 'Organize shared files, folders, logins, brand assets, and access cleanup reminders.', headers);
}

function createCommunicationLogSheet(ss) {
  const headers = ['Log ID', 'Date', 'Project ID', 'Project Name', 'Vendor', 'Channel', 'Topic', 'Summary', 'Follow-Up Needed?', 'Follow-Up Date', 'Notes'];
  createTrackerSheet(ss.getSheetByName('Communication Log'), 'Communication Log', 'Record important vendor conversations, decisions, and follow-up dates.', headers);
}

function createRevisionsSheet(ss) {
  const headers = ['Revision ID', 'Project ID', 'Project Name', 'Vendor', 'Deliverable', 'Revision Round', 'Request Date', 'Due Date', 'Status', 'Approval Date', 'Notes'];
  createTrackerSheet(ss.getSheetByName('Revisions'), 'Revisions', 'Manage revision rounds, requests, due dates, statuses, and final approvals.', headers);
}

function createTrackerSheet(sheet, title, subtitle, headers) {
  formatSheet(sheet, FVD.maxRows, Math.max(headers.length + 2, 16));
  sheet.setFrozenRows(5);
  sheet.setTabColor(FVD.colors.accentBlue);
  sheet.getRange(1, 1, 2, headers.length).merge().setValue(title)
    .setFontSize(22).setFontWeight('bold').setFontColor(FVD.colors.white).setBackground(FVD.colors.darkNavy)
    .setVerticalAlignment('middle');
  sheet.getRange(3, 1, 1, headers.length).merge().setValue(subtitle)
    .setFontColor(FVD.colors.mutedText).setBackground(FVD.colors.white);
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  styleTableHeader(sheet.getRange(5, 1, 1, headers.length));
  sheet.getRange(FVD.firstDataRow, 1, FVD.maxRows - FVD.firstDataRow + 1, headers.length)
    .setBackground(FVD.colors.white)
    .setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle')
    .setWrap(true);
  setColumnWidths(sheet, headers);
}

function createSetupSheet(ss) {
  const sheet = ss.getSheetByName('Setup');
  const listCount = Object.keys(getSetupLists()).length;
  formatSheet(sheet, 120, Math.max(listCount, 22));
  sheet.setTabColor(FVD.colors.nearBlackNavy);
  sheet.getRange(1, 1, 2, listCount).merge().setValue('Setup')
    .setFontSize(22).setFontWeight('bold').setFontColor(FVD.colors.white).setBackground(FVD.colors.darkNavy)
    .setVerticalAlignment('middle');
  sheet.getRange(3, 1, 1, listCount).merge().setValue('Edit these lists to customize dropdown values throughout the workbook.')
    .setFontColor(FVD.colors.mutedText).setBackground(FVD.colors.white);

  const lists = getSetupLists();
  const names = Object.keys(lists);
  names.forEach((name, i) => {
    const col = i + 1;
    sheet.getRange(5, col).setValue(name).setFontWeight('bold').setFontColor(FVD.colors.white).setBackground(FVD.colors.nearBlackNavy);
    sheet.getRange(6, col, lists[name].length, 1).setValues(lists[name].map(v => [v]));
    sheet.getRange(6, col, 80, 1).setBackground(FVD.colors.white).setBorder(true, true, true, true, true, true, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID);
    sheet.setColumnWidth(col, 170);
  });
}

function getSetupLists() {
  return {
    ProjectTypes: ['Website', 'Branding', 'Social Media', 'Copywriting', 'Video', 'Photography', 'Marketing', 'Admin', 'Other'],
    ProjectPriority: ['High', 'Medium', 'Low'],
    ProjectStatus: ['Not Started', 'Active', 'In Progress', 'Waiting on Vendor', 'Waiting on Client', 'On Hold', 'Complete', 'Canceled'],
    VendorRoles: ['Designer', 'Developer', 'Copywriter', 'Editor', 'Virtual Assistant', 'Photographer', 'Video Editor', 'Social Media Manager', 'Bookkeeper', 'Consultant', 'Other'],
    RateTypes: ['Hourly', 'Fixed Project', 'Monthly Retainer', 'Per Deliverable', 'Other'],
    ContractStatus: ['Signed', 'Missing', 'Needed', 'In Review', 'Not Required'],
    Rehire: ['Yes', 'Maybe', 'No'],
    DeliverableStatus: ['Not Started', 'In Progress', 'Submitted', 'Needs Revision', 'Approved', 'Complete', 'Overdue'],
    YesNo: ['Yes', 'No'],
    PaymentTypes: ['Deposit', 'Milestone', 'Final Payment', 'Retainer', 'Reimbursement', 'Bonus', 'Other'],
    PaymentStatus: ['Scheduled', 'Due', 'Overdue', 'Paid', 'Canceled'],
    PaymentMethods: ['ACH', 'Credit Card', 'PayPal', 'Venmo', 'Check', 'Cash', 'Other'],
    DocumentTypes: ['Contract', 'NDA', 'W-9', 'Proposal', 'Scope of Work', 'Invoice', 'Other'],
    FullContractStatus: ['Not Started', 'Needed', 'Requested', 'In Review', 'Signed', 'Missing', 'Not Required'],
    AssetTypes: ['Google Drive Folder', 'Brand Assets', 'Login', 'Photo Folder', 'Copy Doc', 'Design File', 'Website Access', 'Social Media Access', 'Other'],
    AccessLevels: ['View', 'Comment', 'Edit', 'Admin', 'Owner', 'Temporary'],
    RevokeAccess: ['Yes', 'No', 'Not Needed'],
    Channels: ['Email', 'Slack', 'Text', 'Phone', 'Zoom', 'Google Meet', 'In Person', 'Other'],
    RevisionRounds: ['1', '2', '3', '4', '5+'],
    RevisionStatus: ['Requested', 'In Progress', 'Submitted', 'Approved', 'Complete', 'Overdue'],
    Months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
  };
}

function addSampleData(ss) {
  const today = new Date();
  const d = offset => new Date(today.getFullYear(), today.getMonth(), today.getDate() + offset);
  const projects = [
    ['P-001', 'Website Redesign', 'Website', 'Marketing', 'Maya Chen', 'High', 'Active', d(-20), d(14), 5200, 4100, '', 0.72, 'Homepage approved. Product pages in progress.'],
    ['P-002', 'Brand Identity Refresh', 'Branding', 'Brand', 'Sam Patel', 'Medium', 'Waiting on Client', d(-35), d(5), 3500, 2800, '', 0.8, 'Final logo options pending internal feedback.'],
    ['P-003', 'Social Media Management', 'Social Media', 'Marketing', 'Riley Kim', 'Medium', 'In Progress', d(-10), d(20), 1800, 1200, '', 0.45, 'Monthly content calendar underway.'],
    ['P-004', 'Email Welcome Sequence', 'Copywriting', 'Sales', 'Jordan Lee', 'High', 'Waiting on Vendor', d(-12), d(-1), 1200, 1450, '', 0.65, 'Copy draft overdue.'],
    ['P-005', 'Product Photography', 'Photography', 'Products', 'Morgan Wells', 'Low', 'Complete', d(-45), d(-15), 2400, 2350, '', 1, 'Final gallery delivered.'],
    ['P-006', 'YouTube Video Editing', 'Video', 'Content', 'Taylor Brooks', 'High', 'Active', d(-5), d(2), 900, 500, '', 0.55, 'First edit due this week.'],
    ['P-007', 'Blog Content Sprint', 'Copywriting', 'Content', 'Jordan Lee', 'Medium', 'On Hold', d(-22), d(28), 1600, 400, '', 0.25, 'Paused while keywords are finalized.'],
    ['P-008', 'Virtual Assistant Setup', 'Admin', 'Operations', 'Alex Rivera', 'Low', 'Active', d(-7), d(10), 750, 300, '', 0.4, 'Inbox labels and SOP setup started.']
  ];
  ss.getSheetByName('Projects').getRange(6, 1, projects.length, projects[0].length).setValues(projects);

  const vendors = [
    ['V-001', 'Maya Chen', 'Designer', 'maya@example.com', '555-0101', 'mayachen.design', 'Fixed Project', 5200, '50% deposit, 50% final', 'Signed', 4.9, 'Yes', 'Excellent UX and visual polish.'],
    ['V-002', 'Jordan Lee', 'Copywriter', 'jordan@example.com', '555-0102', 'jordanwrites.co', 'Per Deliverable', 400, 'Net 15', 'Needed', 4.6, 'Yes', 'Strong conversion copy.'],
    ['V-003', 'Alex Rivera', 'Virtual Assistant', 'alex@example.com', '555-0103', 'alexops.co', 'Hourly', 35, 'Weekly', 'Missing', 4.7, 'Maybe', 'Great with systems; confirm availability.'],
    ['V-004', 'Sam Patel', 'Designer', 'sam@example.com', '555-0104', 'sampatel.studio', 'Fixed Project', 3500, 'Milestone', 'In Review', 4.8, 'Yes', 'Brand specialist.'],
    ['V-005', 'Taylor Brooks', 'Video Editor', 'taylor@example.com', '555-0105', 'taylorcuts.tv', 'Per Deliverable', 300, 'Due on delivery', 'Signed', 4.4, 'Maybe', 'Fast turnaround.'],
    ['V-006', 'Morgan Wells', 'Photographer', 'morgan@example.com', '555-0106', 'morganwells.photo', 'Fixed Project', 2400, '50% deposit, 50% delivery', 'Signed', 5, 'Yes', 'Beautiful product shots.'],
    ['V-007', 'Riley Kim', 'Social Media Manager', 'riley@example.com', '555-0107', 'rileysocial.co', 'Monthly Retainer', 1800, 'Monthly upfront', 'Not Required', 4.5, 'Yes', 'Keeps content moving.']
  ];
  ss.getSheetByName('Vendors').getRange(6, 1, vendors.length, vendors[0].length).setValues(vendors);

  const deliverables = [
    ['D-001', 'P-001', 'Website Redesign', 'Maya Chen', 'Homepage mockup', d(-3), 'Needs Revision', 'High', 'Yes', '', 'Hero area needs stronger CTA.'],
    ['D-002', 'P-001', 'Website Redesign', 'Maya Chen', 'Product page templates', d(6), 'In Progress', 'High', 'No', '', ''],
    ['D-003', 'P-004', 'Email Welcome Sequence', 'Jordan Lee', '5-email draft sequence', d(-1), 'Overdue', 'High', 'No', '', 'Waiting on first draft.'],
    ['D-004', 'P-006', 'YouTube Video Editing', 'Taylor Brooks', 'First edit cut', d(2), 'In Progress', 'High', 'No', '', ''],
    ['D-005', 'P-005', 'Product Photography', 'Morgan Wells', 'Edited product gallery', d(-16), 'Complete', 'Medium', 'No', d(-15), 'Delivered.'],
    ['D-006', 'P-003', 'Social Media Management', 'Riley Kim', 'June content calendar', d(4), 'Submitted', 'Medium', 'No', '', 'Review captions.'],
    ['D-007', 'P-008', 'Virtual Assistant Setup', 'Alex Rivera', 'Inbox workflow SOP', d(10), 'Not Started', 'Low', 'No', '', ''],
    ['D-008', 'P-002', 'Brand Identity Refresh', 'Sam Patel', 'Final logo files', d(5), 'Submitted', 'Medium', 'No', '', 'Waiting on client pick.']
  ];
  ss.getSheetByName('Deliverables').getRange(6, 1, deliverables.length, deliverables[0].length).setValues(deliverables);

  const payments = [
    ['PAY-001', 'P-001', 'Website Redesign', 'Maya Chen', 'Deposit', 2600, d(-18), d(-18), 'Paid', 'ACH', 'https://example.com/invoice-001', ''],
    ['PAY-002', 'P-001', 'Website Redesign', 'Maya Chen', 'Final Payment', 2600, d(14), '', 'Scheduled', 'ACH', '', ''],
    ['PAY-003', 'P-004', 'Email Welcome Sequence', 'Jordan Lee', 'Milestone', 600, d(-1), '', 'Overdue', 'PayPal', '', 'Pay when draft received.'],
    ['PAY-004', 'P-003', 'Social Media Management', 'Riley Kim', 'Retainer', 1800, d(0), '', 'Due', 'Credit Card', '', ''],
    ['PAY-005', 'P-005', 'Product Photography', 'Morgan Wells', 'Final Payment', 1200, d(-14), d(-13), 'Paid', 'ACH', 'https://example.com/invoice-005', ''],
    ['PAY-006', 'P-006', 'YouTube Video Editing', 'Taylor Brooks', 'Milestone', 300, d(2), '', 'Scheduled', 'Venmo', '', ''],
    ['PAY-007', 'P-002', 'Brand Identity Refresh', 'Sam Patel', 'Final Payment', 1750, d(5), '', 'Scheduled', 'ACH', '', ''],
    ['PAY-008', 'P-008', 'Virtual Assistant Setup', 'Alex Rivera', 'Hourly', 300, d(7), '', 'Scheduled', 'PayPal', '', '']
  ];
  ss.getSheetByName('Payments').getRange(6, 1, payments.length, payments[0].length).setValues(payments);

  const contracts = [
    ['C-001', 'P-001', 'Website Redesign', 'Maya Chen', 'Contract', 'Signed', d(-25), d(-20), d(-20), 'https://example.com/contract-001', ''],
    ['C-002', 'P-004', 'Email Welcome Sequence', 'Jordan Lee', 'Scope of Work', 'Needed', d(-12), d(-2), '', '', 'Scope needs signature.'],
    ['C-003', 'P-008', 'Virtual Assistant Setup', 'Alex Rivera', 'NDA', 'Missing', d(-7), d(0), '', '', 'Access should wait until signed.'],
    ['C-004', 'P-002', 'Brand Identity Refresh', 'Sam Patel', 'Proposal', 'In Review', d(-15), d(3), '', 'https://example.com/proposal-004', ''],
    ['C-005', 'P-005', 'Product Photography', 'Morgan Wells', 'Contract', 'Signed', d(-50), d(-45), d(-44), 'https://example.com/contract-005', ''],
    ['C-006', 'P-003', 'Social Media Management', 'Riley Kim', 'Contract', 'Not Required', '', '', '', '', 'Monthly vendor agreement already on file.']
  ];
  ss.getSheetByName('Contracts').getRange(6, 1, contracts.length, contracts[0].length).setValues(contracts);

  const assets = [
    ['A-001', 'P-001', 'Website Redesign', 'Maya Chen', 'Google Drive Folder', 'Website project folder', 'https://drive.google.com', 'Edit', d(-20), 'No', ''],
    ['A-002', 'P-002', 'Brand Identity Refresh', 'Sam Patel', 'Brand Assets', 'Current logo files', 'https://drive.google.com', 'View', d(-18), 'No', ''],
    ['A-003', 'P-003', 'Social Media Management', 'Riley Kim', 'Social Media Access', 'Instagram scheduler', 'Later.com', 'Admin', d(-10), 'Yes', 'Revoke if retainer ends.'],
    ['A-004', 'P-008', 'Virtual Assistant Setup', 'Alex Rivera', 'Login', 'Helpdesk login', 'Helpdesk app', 'Temporary', d(-3), 'Yes', 'Confirm NDA first.'],
    ['A-005', 'P-005', 'Product Photography', 'Morgan Wells', 'Photo Folder', 'Raw photo upload folder', 'https://drive.google.com', 'Edit', d(-40), 'Not Needed', '']
  ];
  ss.getSheetByName('Assets').getRange(6, 1, assets.length, assets[0].length).setValues(assets);

  const logs = [
    ['L-001', d(-5), 'P-001', 'Website Redesign', 'Maya Chen', 'Zoom', 'Homepage review', 'Reviewed homepage mockup and requested stronger CTA.', 'Yes', d(1), 'Follow up on revised hero.'],
    ['L-002', d(-3), 'P-004', 'Email Welcome Sequence', 'Jordan Lee', 'Email', 'Draft status', 'Asked for delivery update on overdue draft.', 'Yes', d(0), ''],
    ['L-003', d(-2), 'P-003', 'Social Media Management', 'Riley Kim', 'Slack', 'Content calendar', 'Calendar submitted for review.', 'No', '', ''],
    ['L-004', d(-1), 'P-008', 'Virtual Assistant Setup', 'Alex Rivera', 'Google Meet', 'Systems handoff', 'Walked through inbox process and shared SOP outline.', 'Yes', d(5), ''],
    ['L-005', d(-12), 'P-005', 'Product Photography', 'Morgan Wells', 'Email', 'Final gallery', 'Approved final photo gallery.', 'No', '', '']
  ];
  ss.getSheetByName('Communication Log').getRange(6, 1, logs.length, logs[0].length).setValues(logs);

  const revisions = [
    ['R-001', 'P-001', 'Website Redesign', 'Maya Chen', 'Homepage mockup', '1', d(-3), d(1), 'In Progress', '', 'CTA and testimonial section updates.'],
    ['R-002', 'P-004', 'Email Welcome Sequence', 'Jordan Lee', 'Email draft', '1', d(-2), d(-1), 'Overdue', '', 'Need welcome email rewrite.'],
    ['R-003', 'P-006', 'YouTube Video Editing', 'Taylor Brooks', 'First edit cut', '1', d(0), d(4), 'Requested', '', 'Tighten intro pacing.'],
    ['R-004', 'P-005', 'Product Photography', 'Morgan Wells', 'Gallery edits', '2', d(-18), d(-16), 'Approved', d(-15), 'Completed.']
  ];
  ss.getSheetByName('Revisions').getRange(6, 1, revisions.length, revisions[0].length).setValues(revisions);
}

function addDropdowns(ss) {
  const setup = ss.getSheetByName('Setup');
  const ranges = {};
  Object.keys(getSetupLists()).forEach((name, i) => {
    ranges[name] = setup.getRange(6, i + 1, 80, 1);
  });

  addValidation(ss, 'Projects', 6, 3, 'ProjectTypes', ranges);
  addValidation(ss, 'Projects', 6, 6, 'ProjectPriority', ranges);
  addValidation(ss, 'Projects', 6, 7, 'ProjectStatus', ranges);

  addValidation(ss, 'Vendors', 6, 3, 'VendorRoles', ranges);
  addValidation(ss, 'Vendors', 6, 7, 'RateTypes', ranges);
  addValidation(ss, 'Vendors', 6, 10, 'ContractStatus', ranges);
  addValidation(ss, 'Vendors', 6, 12, 'Rehire', ranges);

  addValidation(ss, 'Deliverables', 6, 7, 'DeliverableStatus', ranges);
  addValidation(ss, 'Deliverables', 6, 8, 'ProjectPriority', ranges);
  addValidation(ss, 'Deliverables', 6, 9, 'YesNo', ranges);

  addValidation(ss, 'Payments', 6, 5, 'PaymentTypes', ranges);
  addValidation(ss, 'Payments', 6, 9, 'PaymentStatus', ranges);
  addValidation(ss, 'Payments', 6, 10, 'PaymentMethods', ranges);

  addValidation(ss, 'Contracts', 6, 5, 'DocumentTypes', ranges);
  addValidation(ss, 'Contracts', 6, 6, 'FullContractStatus', ranges);

  addValidation(ss, 'Assets', 6, 5, 'AssetTypes', ranges);
  addValidation(ss, 'Assets', 6, 8, 'AccessLevels', ranges);
  addValidation(ss, 'Assets', 6, 10, 'RevokeAccess', ranges);

  addValidation(ss, 'Communication Log', 6, 6, 'Channels', ranges);
  addValidation(ss, 'Communication Log', 6, 9, 'YesNo', ranges);

  addValidation(ss, 'Revisions', 6, 6, 'RevisionRounds', ranges);
  addValidation(ss, 'Revisions', 6, 9, 'RevisionStatus', ranges);

  const calendar = ss.getSheetByName('Calendar');
  calendar.getRange('C6').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ranges.Months, true).build());
}

function addValidation(ss, sheetName, startRow, col, listName, ranges) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(ranges[listName], true).setAllowInvalid(false).build();
  ss.getSheetByName(sheetName).getRange(startRow, col, FVD.maxRows - startRow + 1, 1).setDataValidation(rule);
}

function addFormulas(ss) {
  const projects = ss.getSheetByName('Projects');
  projects.getRange('L6:L300').setFormulaR1C1('=IF(RC1="","",IF(RC[-2]="","No Budget",IF(RC[-1]>RC[-2],"Over Budget","On Track")))');
  projects.getRange('M6:M300').setNumberFormat('0%');
  projects.getRange('J6:K300').setNumberFormat('$#,##0');
  ss.getSheetByName('Payments').getRange('F6:F300').setNumberFormat('$#,##0');
  ss.getSheetByName('Vendors').getRange('H6:H300').setNumberFormat('$#,##0');

  addDashboardFormulas(ss);
  addCalendarFormulas(ss);
}

function addDashboardFormulas(ss) {
  const sheet = ss.getSheetByName('Dashboard');

  const kpiCells = ['B9', 'F9', 'J9', 'N9', 'B15', 'F15', 'J15', 'N15'];
  const formulas = [
    '=COUNTIFS(Projects!G6:G,"<>Complete",Projects!G6:G,"<>Canceled",Projects!G6:G,"<>On Hold",Projects!A6:A,"<>")',
    '=COUNTIFS(Deliverables!F6:F,"<"&TODAY(),Deliverables!G6:G,"<>Complete",Deliverables!G6:G,"<>Approved",Deliverables!A6:A,"<>")',
    '=COUNTIFS(Payments!G6:G,"<="&TODAY(),Payments!I6:I,"<>Paid",Payments!I6:I,"<>Canceled",Payments!A6:A,"<>")',
    '=SUMIFS(Payments!F6:F,Payments!I6:I,"<>Paid",Payments!I6:I,"<>Canceled")',
    '=IFERROR(SUM(FILTER(Payments!F6:F,Payments!A6:A<>"",MONTH(Payments!G6:G)=MONTH(TODAY()),YEAR(Payments!G6:G)=YEAR(TODAY()),Payments!I6:I<>"Canceled")),0)',
    '=COUNTIFS(Projects!L6:L,"Over Budget",Projects!A6:A,"<>")',
    '=COUNTIFS(Contracts!F6:F,"Missing",Contracts!A6:A,"<>")+COUNTIFS(Contracts!F6:F,"Needed",Contracts!A6:A,"<>")+COUNTIFS(Contracts!F6:F,"Not Started",Contracts!A6:A,"<>")',
    '=IFERROR(ROUND(AVERAGE(FILTER(Vendors!K6:K,Vendors!K6:K<>"")),1)&" / 5","-")'
  ];
  kpiCells.forEach((cell, i) => {
    sheet.getRange(cell).setFormula(formulas[i]);
  });
  sheet.getRange('N9').setNumberFormat('$#,##0');
  sheet.getRange('B15').setNumberFormat('$#,##0');

  sheet.getRange('B53').setFormula(getWatchlistFormula());

  sheet.getRange('J53').setFormula(getUpcomingDeadlinesFormula());

  for (let i = 0; i < 12; i++) {
    const row = 70 + i;
    sheet.getRange(row, 2).setFormula(`=IFERROR(INDEX($AB$3:$AB$14,${i + 1}),"")`);
    sheet.getRange(row, 3).setFormula(`=IFERROR(INDEX($AC$3:$AC$14,${i + 1}),"")`);
    sheet.getRange(row, 7).setFormula(`=IFERROR(INDEX($AE$3:$AE$14,${i + 1}),"")`);
    sheet.getRange(row, 8).setFormula(`=IFERROR(INDEX($AF$3:$AF$14,${i + 1}),"")`);
    sheet.getRange(row, 12).setFormula(`=IFERROR(INDEX($AH$3:$AH$14,${i + 1}),"")`);
    sheet.getRange(row, 13).setFormula(`=IFERROR(INDEX($AI$3:$AI$14,${i + 1}),"")`);
  }

  sheet.getRange('AB2').setFormula('=QUERY(Projects!G6:G,"select G, count(G) where G is not null group by G label G \'Status\', count(G) \'Count\'",0)');
  sheet.getRange('AE2').setFormula('=QUERY(Payments!I6:I,"select I, count(I) where I is not null group by I label I \'Status\', count(I) \'Count\'",0)');
  sheet.getRange('AH2').setFormula('=QUERY(Deliverables!G6:G,"select G, count(G) where G is not null group by G label G \'Status\', count(G) \'Count\'",0)');
  sheet.getRange('AB20').setFormula('=QUERY({ARRAYFORMULA(EOMONTH(Payments!G6:G,0)),Payments!F6:F,Payments!A6:A},"select Col1, sum(Col2) where Col3 is not null group by Col1 order by Col1 label Col1 \'Month\', sum(Col2) \'Spend\'",0)');
  sheet.getRange('AB21:AB42').setNumberFormat('mmm yyyy');
  sheet.getRange('J53:J64').setNumberFormat('mmm d');
  sheet.getRange('F53:F64').setNumberFormat('mmm d');
}

function addCalendarFormulas(ss) {
  const sheet = ss.getSheetByName('Calendar');
  sheet.getRange('AA3').setFormula(getCalendarEventsFormula());
  sheet.getRange('B18').setFormula(getCalendarUpcomingThisWeekFormula());
  sheet.getRange('G18').setFormula(getCalendarOverdueItemsFormula());
  addCalendarMonthlySummaryFormulas(sheet);

  for (let r = 0; r < 6; r++) {
    for (let c = 0; c < 7; c++) {
      const cell = sheet.getRange(9 + r, 2 + c);
      const calendarDate = `DATE($F$6,$C$7,1)-WEEKDAY(DATE($F$6,$C$7,1),1)+1+${r * 7 + c}`;
      cell.setFormula(`=IF(MONTH(${calendarDate})=$C$7,DAY(${calendarDate})&IFERROR(CHAR(10)&TEXTJOIN(CHAR(10),TRUE,FILTER($AE$3:$AE,$AA$3:$AA=${calendarDate})),""),"")`);
    }
  }
}

function addCalendarMonthlySummaryFormulas(sheet) {
  const monthStart = 'DATE($F$6,$C$7,1)';
  const monthEnd = 'EOMONTH(DATE($F$6,$C$7,1),0)+1';
  const countFormula = type => `=COUNTIFS($AA$3:$AA,">="&${monthStart},$AA$3:$AA,"<"&${monthEnd}${type ? `,$AB$3:$AB,"${type}"` : ''})`;
  const types = ['Project', 'Deliverable', 'Payment', 'Contract', 'Revision'];

  sheet.getRange('J9').setFormula(countFormula(''));
  types.forEach((type, i) => {
    sheet.getRange(10 + i, 10).setFormula(countFormula(type));
  });
}

function getCalendarUpcomingThisWeekFormula() {
  return '=IFERROR(SORTN(FILTER({$AA$3:$AA,$AB$3:$AB,$AC$3:$AC,$AD$3:$AD},$AA$3:$AA>=TODAY(),$AA$3:$AA<=TODAY()+7),10,0,1,TRUE),"")';
}

function getCalendarOverdueItemsFormula() {
  return '=IFERROR(SORTN(FILTER({$AA$3:$AA,$AB$3:$AB,$AC$3:$AC,$AD$3:$AD},$AA$3:$AA<TODAY(),$AF$3:$AF<>"Complete",$AF$3:$AF<>"Approved",$AF$3:$AF<>"Paid",$AF$3:$AF<>"Signed",$AF$3:$AF<>"Canceled",$AF$3:$AF<>"Not Required"),10,0,1,TRUE),"")';
}

function getWatchlistFormula() {
  return '=IFERROR(QUERY(VSTACK('
    + 'IFNA(FILTER({Deliverables!H6:H300,Deliverables!E6:E300,ARRAYFORMULA(IF(Deliverables!A6:A300<>"","Deliverable","")),Deliverables!D6:D300,Deliverables!F6:F300,Deliverables!G6:G300,ARRAYFORMULA(IF(Deliverables!A6:A300<>"","Review or follow up",""))},Deliverables!A6:A300<>"",Deliverables!F6:F300<TODAY(),Deliverables!G6:G300<>"Complete",Deliverables!G6:G300<>"Approved"),{"","","","","","",""}),'
    + 'IFNA(FILTER({ARRAYFORMULA(IF(Payments!A6:A300<>"","High","")),Payments!E6:E300,ARRAYFORMULA(IF(Payments!A6:A300<>"","Payment","")),Payments!D6:D300,Payments!G6:G300,Payments!I6:I300,ARRAYFORMULA(IF(Payments!A6:A300<>"","Pay or confirm invoice",""))},Payments!A6:A300<>"",Payments!G6:G300<=TODAY(),Payments!I6:I300<>"Paid",Payments!I6:I300<>"Canceled"),{"","","","","","",""}),'
    + 'IFNA(FILTER({ARRAYFORMULA(IF(Contracts!A6:A300<>"","High","")),Contracts!E6:E300,ARRAYFORMULA(IF(Contracts!A6:A300<>"","Contract","")),Contracts!D6:D300,Contracts!H6:H300,Contracts!F6:F300,ARRAYFORMULA(IF(Contracts!A6:A300<>"","Request or file document",""))},Contracts!A6:A300<>"",REGEXMATCH(Contracts!F6:F300,"Missing|Needed|Not Started")),{"","","","","","",""}),'
    + 'IFNA(FILTER({Projects!F6:F300,Projects!B6:B300,ARRAYFORMULA(IF(Projects!A6:A300<>"","Project","")),Projects!E6:E300,Projects!I6:I300,Projects!L6:L300,ARRAYFORMULA(IF(Projects!A6:A300<>"","Review budget",""))},Projects!A6:A300<>"",Projects!L6:L300="Over Budget"),{"","","","","","",""}),'
    + 'IFNA(FILTER({ARRAYFORMULA(IF(Revisions!A6:A300<>"","High","")),Revisions!E6:E300,ARRAYFORMULA(IF(Revisions!A6:A300<>"","Revision","")),Revisions!D6:D300,Revisions!H6:H300,Revisions!I6:I300,ARRAYFORMULA(IF(Revisions!A6:A300<>"","Review revision status",""))},Revisions!A6:A300<>"",Revisions!H6:H300<=TODAY()+3,Revisions!I6:I300<>"Approved",Revisions!I6:I300<>"Complete"),{"","","","","","",""})'
    + '),"select * where Col1 is not null order by Col5 asc limit 12",0),"No urgent items right now")';
}

function getUpcomingDeadlinesFormula() {
  return '=IFERROR(QUERY(VSTACK('
    + 'IFNA(FILTER({Projects!I6:I300,ARRAYFORMULA(IF(Projects!A6:A300<>"","Project","")),Projects!B6:B300,Projects!E6:E300,ARRAYFORMULA(IF(Projects!A6:A300<>"","Due/Launch Date","")),Projects!G6:G300},Projects!A6:A300<>"",Projects!I6:I300>=TODAY(),Projects!I6:I300<=TODAY()+30),{"","","","","",""}),'
    + 'IFNA(FILTER({Deliverables!F6:F300,ARRAYFORMULA(IF(Deliverables!A6:A300<>"","Deliverable","")),Deliverables!C6:C300,Deliverables!D6:D300,Deliverables!E6:E300,Deliverables!G6:G300},Deliverables!A6:A300<>"",Deliverables!F6:F300>=TODAY(),Deliverables!F6:F300<=TODAY()+30),{"","","","","",""}),'
    + 'IFNA(FILTER({Payments!G6:G300,ARRAYFORMULA(IF(Payments!A6:A300<>"","Payment","")),Payments!C6:C300,Payments!D6:D300,Payments!E6:E300,Payments!I6:I300},Payments!A6:A300<>"",Payments!G6:G300>=TODAY(),Payments!G6:G300<=TODAY()+30),{"","","","","",""}),'
    + 'IFNA(FILTER({Contracts!H6:H300,ARRAYFORMULA(IF(Contracts!A6:A300<>"","Contract","")),Contracts!C6:C300,Contracts!D6:D300,Contracts!E6:E300,Contracts!F6:F300},Contracts!A6:A300<>"",Contracts!H6:H300>=TODAY(),Contracts!H6:H300<=TODAY()+30),{"","","","","",""}),'
    + 'IFNA(FILTER({Revisions!H6:H300,ARRAYFORMULA(IF(Revisions!A6:A300<>"","Revision","")),Revisions!C6:C300,Revisions!D6:D300,Revisions!E6:E300,Revisions!I6:I300},Revisions!A6:A300<>"",Revisions!H6:H300>=TODAY(),Revisions!H6:H300<=TODAY()+30),{"","","","","",""})'
    + '),"select * where Col1 is not null order by Col1 asc limit 12",0),"No upcoming deadlines")';
}

function getCalendarEventsFormula() {
  return '=IFERROR(QUERY(VSTACK('
    + 'IFNA(FILTER({Projects!I6:I300,ARRAYFORMULA(IF(Projects!A6:A300<>"","Project","")),Projects!B6:B300,Projects!E6:E300,ARRAYFORMULA(IF(Projects!A6:A300<>"","[P] "&LEFT(Projects!B6:B300,35),"")),Projects!G6:G300},Projects!A6:A300<>"",Projects!I6:I300<>""),{"","","","","",""}),'
    + 'IFNA(FILTER({Deliverables!F6:F300,ARRAYFORMULA(IF(Deliverables!A6:A300<>"","Deliverable","")),Deliverables!E6:E300,Deliverables!D6:D300,ARRAYFORMULA(IF(Deliverables!A6:A300<>"","[D] "&LEFT(Deliverables!E6:E300,35),"")),Deliverables!G6:G300},Deliverables!A6:A300<>"",Deliverables!F6:F300<>""),{"","","","","",""}),'
    + 'IFNA(FILTER({Payments!G6:G300,ARRAYFORMULA(IF(Payments!A6:A300<>"","Payment","")),Payments!E6:E300,Payments!D6:D300,ARRAYFORMULA(IF(Payments!A6:A300<>"","[$] "&LEFT(Payments!E6:E300,35),"")),Payments!I6:I300},Payments!A6:A300<>"",Payments!G6:G300<>""),{"","","","","",""}),'
    + 'IFNA(FILTER({Contracts!H6:H300,ARRAYFORMULA(IF(Contracts!A6:A300<>"","Contract","")),Contracts!E6:E300,Contracts!D6:D300,ARRAYFORMULA(IF(Contracts!A6:A300<>"","[C] "&LEFT(Contracts!E6:E300,35),"")),Contracts!F6:F300},Contracts!A6:A300<>"",Contracts!H6:H300<>""),{"","","","","",""}),'
    + 'IFNA(FILTER({Revisions!H6:H300,ARRAYFORMULA(IF(Revisions!A6:A300<>"","Revision","")),Revisions!E6:E300,Revisions!D6:D300,ARRAYFORMULA(IF(Revisions!A6:A300<>"","[R] "&LEFT(Revisions!E6:E300,35),"")),Revisions!I6:I300},Revisions!A6:A300<>"",Revisions!H6:H300<>""),{"","","","","",""})'
    + '),"select * where Col1 is not null order by Col1 asc",0),"")';
}

function addConditionalFormatting(ss) {
  FVD.sheets.forEach(name => ss.getSheetByName(name).clearConditionalFormatRules());
  addStatusRules(ss.getSheetByName('Projects'), 'G6:G300', ['Complete'], ['Canceled'], ['Waiting on Vendor', 'Waiting on Client', 'On Hold']);
  addTextRule(ss.getSheetByName('Projects'), 'F6:F300', 'High', FVD.colors.softRed);
  addTextRule(ss.getSheetByName('Projects'), 'L6:L300', 'Over Budget', FVD.colors.softRed);
  addDateRules(ss.getSheetByName('Projects'), 'I6:I300');

  addStatusRules(ss.getSheetByName('Deliverables'), 'G6:G300', ['Approved', 'Complete'], ['Overdue', 'Needs Revision'], ['Submitted']);
  addTextRule(ss.getSheetByName('Deliverables'), 'H6:H300', 'High', FVD.colors.softRed);
  addDateRules(ss.getSheetByName('Deliverables'), 'F6:F300');

  addStatusRules(ss.getSheetByName('Payments'), 'I6:I300', ['Paid'], ['Overdue'], ['Due', 'Scheduled']);
  addDateRules(ss.getSheetByName('Payments'), 'G6:G300');

  addStatusRules(ss.getSheetByName('Contracts'), 'F6:F300', ['Signed', 'Not Required'], ['Missing', 'Needed', 'Not Started'], ['Requested', 'In Review']);
  addDateRules(ss.getSheetByName('Contracts'), 'H6:H300');

  addStatusRules(ss.getSheetByName('Revisions'), 'I6:I300', ['Approved', 'Complete'], ['Overdue'], ['Requested', 'In Progress', 'Submitted']);
  addDateRules(ss.getSheetByName('Revisions'), 'H6:H300');

  addStatusRules(ss.getSheetByName('Vendors'), 'J6:J300', ['Signed', 'Not Required'], ['Missing', 'Needed'], ['In Review']);
  addTextRule(ss.getSheetByName('Assets'), 'J6:J300', 'Yes', FVD.colors.softYellow);
  addTextRule(ss.getSheetByName('Communication Log'), 'I6:I300', 'Yes', FVD.colors.softYellow);
  addDateRules(ss.getSheetByName('Communication Log'), 'J6:J300');

  addCalendarConditionalFormatting(ss.getSheetByName('Calendar'));
}

function addStatusRules(sheet, a1, positive, negative, waiting) {
  positive.forEach(v => addTextRule(sheet, a1, v, FVD.colors.softGreen));
  negative.forEach(v => addTextRule(sheet, a1, v, FVD.colors.softRed));
  waiting.forEach(v => addTextRule(sheet, a1, v, FVD.colors.softYellow));
}

function addTextRule(sheet, a1, text, color) {
  const range = sheet.getRange(a1);
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(color)
    .setRanges([range])
    .build());
  sheet.setConditionalFormatRules(rules);
}

function addDateRules(sheet, a1) {
  const range = sheet.getRange(a1);
  const firstCell = range.getCell(1, 1).getA1Notation();
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(${firstCell}<>"",${firstCell}<TODAY())`)
    .setBackground(FVD.colors.softRed)
    .setRanges([range])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(${firstCell}<>"",${firstCell}>=TODAY(),${firstCell}<=TODAY()+7)`)
    .setBackground(FVD.colors.softYellow)
    .setRanges([range])
    .build());
  sheet.setConditionalFormatRules(rules);
}

function addCalendarConditionalFormatting(sheet) {
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(B9<>"",DATE($F$6,$C$7,VALUE(REGEXEXTRACT(B9,"^\\d+")))=TODAY())')
    .setBackground('#EAF8F6')
    .setFontColor(FVD.colors.darkNavy)
    .setRanges([sheet.getRange('B9:H14')])
    .build());

  sheet.setConditionalFormatRules(rules);
}

function addDashboardCharts(ss) {
  const sheet = ss.getSheetByName('Dashboard');
  const existing = sheet.getCharts();
  existing.forEach(chart => sheet.removeChart(chart));

  const palette = [FVD.colors.accentBlue, FVD.colors.softTeal, FVD.colors.softCoral, FVD.colors.softBeige, FVD.colors.darkNavy, '#9CA3AF'];
  const chartConfigs = [
    { title: 'Project Status Breakdown', range: 'AB2:AC14', row: 23, col: 2, type: Charts.ChartType.PIE },
    { title: 'Payment Status Breakdown', range: 'AE2:AF14', row: 23, col: 10, type: Charts.ChartType.PIE },
    { title: 'Monthly Vendor Spend', range: 'AB20:AC34', row: 37, col: 2, type: Charts.ChartType.COLUMN },
    { title: 'Deliverables by Status', range: 'AH2:AI14', row: 37, col: 10, type: Charts.ChartType.BAR }
  ];
  const chartWidth = 560;
  const panelWidth = 8 * 92;
  const centeredOffset = Math.floor((panelWidth - chartWidth) / 2);

  chartConfigs.forEach(cfg => {
    let builder = sheet.newChart()
      .addRange(sheet.getRange(cfg.range))
      .setChartType(cfg.type)
      .setPosition(cfg.row, cfg.col, centeredOffset, 6)
      .setOption('title', cfg.title)
      .setOption('titleTextStyle', { color: FVD.colors.textDark, fontSize: 13, bold: true })
      .setOption('backgroundColor', FVD.colors.white)
      .setOption('colors', palette)
      .setOption('legend', { position: 'right', textStyle: { color: FVD.colors.mutedText, fontSize: 10 } })
      .setOption('width', chartWidth)
      .setOption('height', 255);
    if (cfg.type === Charts.ChartType.COLUMN || cfg.type === Charts.ChartType.BAR) {
      builder = builder
        .setOption('hAxis', { textStyle: { color: FVD.colors.mutedText } })
        .setOption('vAxis', { textStyle: { color: FVD.colors.mutedText }, minValue: 0 });
    }
    sheet.insertChart(builder.build());
  });
}

function finalizeWorkbook(ss) {
  FVD.sheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    sheet.setHiddenGridlines(true);
    sheet.getDataRange().setFontFamily(FVD.font);
  });
  ss.setActiveSheet(ss.getSheetByName('Start Here'));
}

function formatSheet(sheet, rows, cols) {
  sheet.clear();
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < cols) sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
  if (sheet.getMaxRows() > rows) sheet.deleteRows(rows + 1, sheet.getMaxRows() - rows);
  if (sheet.getMaxColumns() > cols) sheet.deleteColumns(cols + 1, sheet.getMaxColumns() - cols);
  sheet.getRange(1, 1, rows, cols).setBackground(FVD.colors.lightGray).setFontFamily(FVD.font).setFontColor(FVD.colors.textDark);
  sheet.setRowHeights(1, rows, 28);
}

function styleHeader(range) {
  range.setBackground(FVD.colors.darkNavy)
    .setFontColor(FVD.colors.white)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

function styleTableHeader(range) {
  styleHeader(range);
  range.setWrap(true)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, FVD.colors.darkNavy, SpreadsheetApp.BorderStyle.SOLID);
}

function sectionTitle(sheet, a1, title) {
  sheet.getRange(a1).merge().setValue(title)
    .setFontSize(13).setFontWeight('bold').setFontColor(FVD.colors.white)
    .setBackground(FVD.colors.nearBlackNavy)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, FVD.colors.nearBlackNavy, SpreadsheetApp.BorderStyle.SOLID);
}

function stylePanel(sheet, a1) {
  sheet.getRange(a1)
    .setBackground(FVD.colors.white)
    .setBorder(true, true, true, true, null, null, FVD.colors.borderGray, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function setColumnWidths(sheet, headers) {
  const widths = headers.map(header => {
    if (/Notes|Summary/.test(header)) return 230;
    if (/URL|Link|Website|Portfolio|Location|Email/.test(header)) return 180;
    if (/Date|Due|Paid|Requested|Received|Shared|Approval/.test(header)) return 115;
    if (/Amount|Budget|Spend|Rate|Progress/.test(header)) return 110;
    if (/ID/.test(header)) return 95;
    return 145;
  });
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
}