/**
 * ============================================================================
 * Winton Wallet
 * Register.gs
 * ----------------------------------------------------------------------------
 * Register management and transaction operations.
 * ============================================================================
 */

/**
 * Register column definitions.
 */
const WW_REGISTER = Object.freeze({

  COL: Object.freeze({
    DATE: 1,
    ACCOUNT: 2,
    DESCRIPTION: 3,
    CATEGORY: 4,
    PAYMENT_METHOD: 5,
    DEPOSIT: 6,
    WITHDRAWAL: 7,
    BALANCE: 8,
    SOURCE: 9,
    NOTES: 10
  })

});

/**
 * Returns the Register sheet.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function wwGetRegisterSheet() {

  return SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.REGISTER);

}

/**
 * Returns the next available row in the Register.
 *
 * @returns {number}
 */
function wwGetNextRegisterRow() {

  const sheet = wwGetRegisterSheet();

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 2;
  }

  return lastRow + 1;

}

/**
 * Determines whether a register row is empty.
 *
 * @param {number} row
 * @returns {boolean}
 */
function wwIsRegisterRowEmpty(row) {

  const sheet = wwGetRegisterSheet();

  const values = sheet
    .getRange(row, 1, 1, 10)
    .getValues()[0];

  return values.every(value => value === "");

}
/**
 * Adds a transaction to the Register.
 *
 * @param {Object} transaction
 */
function wwAddTransaction(transaction) {

  const sheet = wwGetRegisterSheet();
  const row = wwGetNextRegisterRow();

  sheet.getRange(row, WW_REGISTER.COL.DATE)
    .setValue(transaction.date);

  sheet.getRange(row, WW_REGISTER.COL.ACCOUNT)
    .setValue(transaction.account);

  sheet.getRange(row, WW_REGISTER.COL.DESCRIPTION)
    .setValue(transaction.description);

  sheet.getRange(row, WW_REGISTER.COL.CATEGORY)
    .setValue(transaction.category);

  sheet.getRange(row, WW_REGISTER.COL.PAYMENT_METHOD)
    .setValue(transaction.paymentMethod);

  sheet.getRange(row, WW_REGISTER.COL.DEPOSIT)
    .setValue(transaction.deposit || "");

  sheet.getRange(row, WW_REGISTER.COL.WITHDRAWAL)
    .setValue(transaction.withdrawal || "");

  sheet.getRange(row, WW_REGISTER.COL.SOURCE)
    .setValue(transaction.source || "");

  sheet.getRange(row, WW_REGISTER.COL.NOTES)
    .setValue(transaction.notes || "");

}

/**
 * Adds a deposit transaction.
 *
 * @param {Date} date
 * @param {string} account
 * @param {string} description
 * @param {string} category
 * @param {string} paymentMethod
 * @param {number} amount
 * @param {string} source
 * @param {string} notes
 */
function wwAddDeposit(
  date,
  account,
  description,
  category,
  paymentMethod,
  amount,
  source,
  notes
) {

  wwAddTransaction({
    date: date,
    account: account,
    description: description,
    category: category,
    paymentMethod: paymentMethod,
    deposit: amount,
    withdrawal: "",
    source: source,
    notes: notes
  });

}

/**
 * Adds a withdrawal transaction.
 *
 * @param {Date} date
 * @param {string} account
 * @param {string} description
 * @param {string} category
 * @param {string} paymentMethod
 * @param {number} amount
 * @param {string} source
 * @param {string} notes
 */
function wwAddWithdrawal(
  date,
  account,
  description,
  category,
  paymentMethod,
  amount,
  source,
  notes
) {

  wwAddTransaction({
    date: date,
    account: account,
    description: description,
    category: category,
    paymentMethod: paymentMethod,
    deposit: "",
    withdrawal: amount,
    source: source,
    notes: notes
  });

}
/**
 * Recalculates the running balance for every transaction
 * in the Register.
 */
function wwRecalculateRegisterBalances() {

  const sheet = wwGetRegisterSheet();

  const lastRow = sheet.getLastRow();

  // Nothing to calculate if there are no transactions.
  if (lastRow < 2) {
    return;
  }

  const numRows = lastRow - 1;

  const values = sheet
    .getRange(2, 1, numRows, 10)
    .getValues();

  let runningBalance = 0;

  values.forEach((row, index) => {

    const deposit = Number(row[WW_REGISTER.COL.DEPOSIT - 1]) || 0;
    const withdrawal = Number(row[WW_REGISTER.COL.WITHDRAWAL - 1]) || 0;

    runningBalance += deposit;
    runningBalance -= withdrawal;

    sheet
      .getRange(index + 2, WW_REGISTER.COL.BALANCE)
      .setValue(runningBalance);

  });

}

/**
 * Returns the current register balance.
 *
 * @returns {number}
 */
function wwGetCurrentBalance() {

  const sheet = wwGetRegisterSheet();

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 0;
  }

  return Number(
    sheet
      .getRange(lastRow, WW_REGISTER.COL.BALANCE)
      .getValue()
  ) || 0;

}

/**
 * Adds a transaction and immediately updates balances.
 *
 * @param {Object} transaction
 */
function wwPostTransaction(transaction) {

  wwAddTransaction(transaction);

  wwRecalculateRegisterBalances();

}

/**
 * Clears all running balances and recalculates them.
 */
function wwResetBalances() {

  const sheet = wwGetRegisterSheet();

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  sheet
    .getRange(
      2,
      WW_REGISTER.COL.BALANCE,
      lastRow - 1,
      1
    )
    .clearContent();

  wwRecalculateRegisterBalances();

}
/**
 * Returns a transaction object for the specified register row.
 *
 * @param {number} row
 * @returns {Object|null}
 */
function wwGetTransaction(row) {

  const sheet = wwGetRegisterSheet();

  if (row < 2 || row > sheet.getLastRow()) {
    return null;
  }

  const values = sheet
    .getRange(row, 1, 1, 10)
    .getValues()[0];

  return {
    row: row,
    date: values[WW_REGISTER.COL.DATE - 1],
    account: values[WW_REGISTER.COL.ACCOUNT - 1],
    description: values[WW_REGISTER.COL.DESCRIPTION - 1],
    category: values[WW_REGISTER.COL.CATEGORY - 1],
    paymentMethod: values[WW_REGISTER.COL.PAYMENT_METHOD - 1],
    deposit: values[WW_REGISTER.COL.DEPOSIT - 1],
    withdrawal: values[WW_REGISTER.COL.WITHDRAWAL - 1],
    balance: values[WW_REGISTER.COL.BALANCE - 1],
    source: values[WW_REGISTER.COL.SOURCE - 1],
    notes: values[WW_REGISTER.COL.NOTES - 1]
  };

}

/**
 * Returns the last transaction row.
 *
 * @returns {number}
 */
function wwGetLastTransactionRow() {

  const sheet = wwGetRegisterSheet();

  return Math.max(sheet.getLastRow(), 2);

}

/**
 * Returns the number of transactions.
 *
 * @returns {number}
 */
function wwGetTransactionCount() {

  const sheet = wwGetRegisterSheet();

  return Math.max(sheet.getLastRow() - 1, 0);

}

/**
 * Determines whether the Register contains transactions.
 *
 * @returns {boolean}
 */
function wwHasTransactions() {

  return wwGetTransactionCount() > 0;

}

/**
 * Validates a transaction object.
 *
 * @param {Object} transaction
 */
function wwValidateTransaction(transaction) {

  if (!transaction) {
    throw new Error("Transaction is required.");
  }

  if (!transaction.date) {
    throw new Error("Transaction date is required.");
  }

  if (!transaction.account) {
    throw new Error("Account is required.");
  }

  if (!transaction.description) {
    throw new Error("Description is required.");
  }

  const deposit = Number(transaction.deposit) || 0;
  const withdrawal = Number(transaction.withdrawal) || 0;

  if (deposit > 0 && withdrawal > 0) {
    throw new Error("A transaction cannot contain both a deposit and a withdrawal.");
  }

  if (deposit === 0 && withdrawal === 0) {
    throw new Error("Enter either a deposit or a withdrawal.");
  }

}

/**
 * Posts a validated transaction.
 *
 * @param {Object} transaction
 */
function wwPostValidatedTransaction(transaction) {

  wwValidateTransaction(transaction);

  wwPostTransaction(transaction);

}
