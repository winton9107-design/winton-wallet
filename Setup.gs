/**
 * ============================================================================
 * Winton Wallet
 * Setup.gs
 * ----------------------------------------------------------------------------
 * Creates and formats the Winton Wallet workbook.
 * ============================================================================
 */

/**
 * Creates or rebuilds the workbook.
 */
function wwSetupWorkbook() {

  const ss = SpreadsheetApp.getActive();

  const sheetNames = [
    WW.SHEETS.DASHBOARD,
    WW.SHEETS.REGISTER,
    WW.SHEETS.BILLS,
    WW.SHEETS.INCOME,
    WW.SHEETS.ACCOUNTS,
    WW.SHEETS.REPORTS,
    WW.SHEETS.SETTINGS,
    WW.SHEETS.LISTS
  ];

  sheetNames.forEach(name => {

    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    sheet.clear();
    sheet.clearFormats();

  });

  wwSetupDashboard();
  wwSetupRegister();
  wwSetupBills();
  wwSetupIncome();
  wwSetupAccounts();
  wwSetupReports();
  wwSetupSettings();
  wwSetupLists();

  SpreadsheetApp.flush();

}
/**
 * Creates the Dashboard sheet.
 */
function wwSetupDashboard() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.DASHBOARD);

  sheet.setHiddenGridlines(true);

  sheet.setColumnWidths(1, 8, 160);

  sheet.setRowHeight(1, 40);

  sheet.getRange("A1:H1").merge();

  sheet.getRange("A1")
    .setValue("Winton Wallet Dashboard");

  wwFormatTitle(sheet.getRange("A1"));

}
/**
 * Creates the Register sheet.
 */
function wwSetupRegister() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.REGISTER);

  sheet.setHiddenGridlines(true);

  const headers = [
    "Date",
    "Account",
    "Description",
    "Category",
    "Payment Method",
    "Deposit",
    "Withdrawal",
    "Running Balance",
    "Source",
    "Notes"
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  wwFormatHeader(sheet.getRange(1, 1, 1, headers.length));

  const widths = [100, 120, 220, 150, 140, 100, 100, 120, 120, 220];

  widths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });

  sheet.setFrozenRows(1);

}
/**
 * Creates the Scheduled Bills sheet.
 */
function wwSetupBills() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.BILLS);

  sheet.setHiddenGridlines(true);

  const headers = [
    "Bill Name",
    "Payee",
    "Amount",
    "Due Date",
    "Repeat",
    "Next Due",
    "Payment Method",
    "Account",
    "Category",
    "Active"
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  wwFormatHeader(sheet.getRange(1, 1, 1, headers.length));

  const widths = [180, 180, 100, 110, 120, 110, 140, 120, 150, 80];

  widths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });

  sheet.setFrozenRows(1);

}
/**
 * Creates the Income sheet.
 */
function wwSetupIncome() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.INCOME);

  sheet.setHiddenGridlines(true);

  const headers = [
    "Date",
    "Description",
    "Amount",
    "Account",
    "Notes"
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  wwFormatHeader(sheet.getRange(1, 1, 1, headers.length));

  const widths = [100, 250, 120, 120, 250];

  widths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });

  sheet.setFrozenRows(1);

}

/**
 * Creates the Accounts sheet.
 */
function wwSetupAccounts() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.ACCOUNTS);

  sheet.setHiddenGridlines(true);

  const headers = [
    "Account",
    "Current Balance",
    "Last Updated"
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  wwFormatHeader(sheet.getRange(1, 1, 1, headers.length));

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 160);

  sheet.setFrozenRows(1);

}

/**
 * Creates the Reports sheet.
 */
function wwSetupReports() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.REPORTS);

  sheet.setHiddenGridlines(true);

  sheet.getRange("A1")
    .setValue("Reports");

  wwFormatTitle(sheet.getRange("A1"));

}

/**
 * Creates the Settings sheet.
 */
function wwSetupSettings() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.SETTINGS);

  sheet.setHiddenGridlines(true);

  const headers = [
    "Setting",
    "Value"
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  wwFormatHeader(sheet.getRange(1, 1, 1, headers.length));

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 250);

  sheet.setFrozenRows(1);

}

/**
 * Creates the Lists sheet.
 */
function wwSetupLists() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.LISTS);

  sheet.setHiddenGridlines(true);

  sheet.getRange("A1")
    .setValue("Reference Lists");

  wwFormatTitle(sheet.getRange("A1"));

}
/**
 * Initializes default data after the workbook has been created.
 */
function wwInitializeDefaults() {

  const ss = SpreadsheetApp.getActive();

  // Accounts
  const accountsSheet = ss.getSheetByName(WW.SHEETS.ACCOUNTS);

  accountsSheet.getRange("A2:C3").setValues([
    ["Checking", 0.00, ""],
    ["Savings", 0.00, ""]
  ]);

  // Lists
  const listsSheet = ss.getSheetByName(WW.SHEETS.LISTS);

  listsSheet.clearContents();

  listsSheet.getRange("A1").setValue("Accounts");
  listsSheet.getRange("A2:A3").setValues([
    ["Checking"],
    ["Savings"]
  ]);

  listsSheet.getRange("C1").setValue("Payment Methods");
  listsSheet.getRange("C2:C5").setValues([
    ["ACH"],
    ["Check"],
    ["Debit Card"],
    ["Transfer"]
  ]);

  listsSheet.getRange("E1").setValue("Repeat Patterns");
  listsSheet.getRange("E2:E6").setValues([
    ["One Time"],
    ["Weekly"],
    ["Bi-Weekly"],
    ["Monthly"],
    ["Yearly"]
  ]);

  listsSheet.getRange("G1").setValue("Categories");
  listsSheet.getRange("G2:G8").setValues([
    ["Housing"],
    ["Utilities"],
    ["Transportation"],
    ["Food"],
    ["Insurance"],
    ["Entertainment"],
    ["Other"]
  ]);

}

/**
 * Public entry point used by the menu to build the workbook.
 */
function wwInstall() {

  wwSetupWorkbook();
  wwInitializeDefaults();

  Logger.log(WW.APP.NAME + " installed successfully.");

}
