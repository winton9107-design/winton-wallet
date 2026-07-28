/**
 * ============================================================================
 * Winton Wallet
 * Menu.gs
 * ----------------------------------------------------------------------------
 * Creates the custom Winton Wallet menu.
 * ============================================================================
 */

/**
 * Runs automatically when the spreadsheet opens.
 */
function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu("💜 Winton Wallet")
    .addItem("Setup Workbook", "wwInstall")
    .addSeparator()
    .addItem("Refresh Dashboard", "wwRefreshDashboard")
    .addItem("Post Due Bills", "wwPostDueBills")
    .addItem("Recalculate Balances", "wwRecalculateBalances")
    .addSeparator()
    .addItem("Settings", "wwOpenSettings")
    .addSeparator()
    .addItem("About Winton Wallet", "wwAbout")
    .addToUi();

}

/**
 * Displays information about the application.
 */
function wwAbout() {

  SpreadsheetApp.getUi().alert(
    `${WW.APP.NAME}\n\nVersion ${WW.APP.VERSION}\n\nCreated for Katie Winton`
  );

}

/**
 * Placeholder until Dashboard.gs is implemented.
 */
function wwRefreshDashboard() {

  SpreadsheetApp.getUi().alert(
    "Dashboard refresh will be available after Dashboard.gs is installed."
  );

}

/**
 * Placeholder until Posting.gs is implemented.
 */
function wwPostDueBills() {

  const posted = wwProcessDueBills();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    posted + " bill(s) posted.",
    "Winton Wallet",
    5
  );

}
/**
 * Placeholder until Register.gs is implemented.
 */
function wwRecalculateBalances() {

  SpreadsheetApp.getUi().alert(
    "Balance recalculation will be available after Register.gs is installed."
  );

}

/**
 * Placeholder until Settings.gs is implemented.
 */
function wwOpenSettings() {

  SpreadsheetApp.getUi().alert(
    "Settings will be available after Settings.gs is installed."
  );

}
