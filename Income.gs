/**
 * ============================================================================
 * Winton Wallet
 * Income.gs
 * ----------------------------------------------------------------------------
 * Scheduled Income Management
 * ============================================================================
 */

/**
 * Scheduled Income column definitions.
 */
const WW_INCOME = Object.freeze({

  COL: Object.freeze({
    ACTIVE: 1,
    NAME: 2,
    ACCOUNT: 3,
    CATEGORY: 4,
    PAYMENT_METHOD: 5,
    AMOUNT: 6,
    PAY_DATE: 7,
    REPEAT: 8,
    SOURCE: 9,
    NOTES: 10,
    LAST_RECEIVED: 11,
    NEXT_PAY: 12
  })

});

/**
 * Returns the Income sheet.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function wwGetIncomeSheet() {

  return SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.INCOME);

}

/**
 * Returns the last income row.
 *
 * @returns {number}
 */
function wwGetLastIncomeRow() {

  const sheet = wwGetIncomeSheet();

  return Math.max(sheet.getLastRow(), 2);

}

/**
 * Returns the number of income records.
 *
 * @returns {number}
 */
function wwGetIncomeCount() {

  return Math.max(wwGetLastIncomeRow() - 1, 0);

}

/**
 * Returns true if Income contains records.
 *
 * @returns {boolean}
 */
function wwHasIncome() {

  return wwGetIncomeCount() > 0;

}
/**
 * Returns an income record from the specified row.
 *
 * @param {number} row
 * @returns {Object|null}
 */
function wwGetIncome(row) {

  const sheet = wwGetIncomeSheet();

  if (row < 2 || row > sheet.getLastRow()) {
    return null;
  }

  const values = sheet
    .getRange(row, 1, 1, 12)
    .getValues()[0];

  return {
    row: row,
    active: values[WW_INCOME.COL.ACTIVE - 1],
    name: values[WW_INCOME.COL.NAME - 1],
    account: values[WW_INCOME.COL.ACCOUNT - 1],
    category: values[WW_INCOME.COL.CATEGORY - 1],
    paymentMethod: values[WW_INCOME.COL.PAYMENT_METHOD - 1],
    amount: values[WW_INCOME.COL.AMOUNT - 1],
    payDate: values[WW_INCOME.COL.PAY_DATE - 1],
    repeat: values[WW_INCOME.COL.REPEAT - 1],
    source: values[WW_INCOME.COL.SOURCE - 1],
    notes: values[WW_INCOME.COL.NOTES - 1],
    lastReceived: values[WW_INCOME.COL.LAST_RECEIVED - 1],
    nextPay: values[WW_INCOME.COL.NEXT_PAY - 1]
  };

}

/**
 * Returns every scheduled income record.
 *
 * @returns {Object[]}
 */
function wwGetIncomeRecords() {

  const records = [];

  const lastRow = wwGetLastIncomeRow();

  for (let row = 2; row <= lastRow; row++) {

    const income = wwGetIncome(row);

    if (income) {
      records.push(income);
    }

  }

  return records;

}

/**
 * Returns only active income records.
 *
 * @returns {Object[]}
 */
function wwGetActiveIncome() {

  return wwGetIncomeRecords().filter(income => {

    const value = String(income.active).toLowerCase();

    return value === "true" ||
           value === "yes" ||
           value === "y" ||
           value === "active";

  });

}

/**
 * Returns all income records due today or earlier.
 *
 * @returns {Object[]}
 */
function wwGetDueIncome() {

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return wwGetActiveIncome().filter(income => {

    if (!(income.nextPay instanceof Date)) {
      return false;
    }

    const payDate = new Date(income.nextPay);
    payDate.setHours(0, 0, 0, 0);

    return payDate <= today;

  });

}
/**
 * Posts a scheduled income record to the Register.
 *
 * @param {Object} income
 */
function wwPostIncome(income) {

  wwPostValidatedTransaction({
    date: new Date(),
    account: income.account,
    description: income.name,
    category: income.category,
    paymentMethod: income.paymentMethod,
    deposit: income.amount,
    withdrawal: "",
    source: income.source,
    notes: income.notes
  });

  wwUpdateIncomeAfterPosting(income);

}

/**
 * Updates an income record after it has been posted.
 *
 * @param {Object} income
 */
function wwUpdateIncomeAfterPosting(income) {

  const sheet = wwGetIncomeSheet();
  const today = new Date();

  sheet
    .getRange(income.row, WW_INCOME.COL.LAST_RECEIVED)
    .setValue(today);

  sheet
    .getRange(income.row, WW_INCOME.COL.NEXT_PAY)
    .setValue(
      wwCalculateNextPayDate(today, income.repeat)
    );

}

/**
 * Calculates the next pay date based on the income's
 * repeat schedule.
 *
 * Supported:
 * Once
 * Weekly
 * Bi-Weekly
 * Monthly
 * Yearly
 *
 * @param {Date} date
 * @param {string} repeat
 * @returns {Date|string}
 */
function wwCalculateNextPayDate(date, repeat) {

  const next = new Date(date);

  switch (String(repeat).toLowerCase()) {

    case "weekly":
      next.setDate(next.getDate() + 7);
      break;

    case "bi-weekly":
    case "biweekly":
      next.setDate(next.getDate() + 14);
      break;

    case "monthly":
      next.setMonth(next.getMonth() + 1);
      break;

    case "yearly":
      next.setFullYear(next.getFullYear() + 1);
      break;

    case "once":
    default:
      return "";

  }

  return next;

}

/**
 * Processes every income record that is currently due.
 *
 * @returns {number}
 */
function wwProcessDueIncome() {

  const dueIncome = wwGetDueIncome();

  dueIncome.forEach(income => {

    wwPostIncome(income);

  });

  return dueIncome.length;

}
/**
 * Validates a scheduled income record.
 *
 * @param {Object} income
 */
function wwValidateIncome(income) {

  if (!income) {
    throw new Error("Income record is required.");
  }

  if (!income.name) {
    throw new Error("Income name is required.");
  }

  if (!income.account) {
    throw new Error("Account is required.");
  }

  if (!income.amount || Number(income.amount) <= 0) {
    throw new Error("Income amount must be greater than zero.");
  }

  if (!(income.nextPay instanceof Date)) {
    throw new Error("A valid Next Pay date is required.");
  }

}

/**
 * Posts a validated income record.
 *
 * @param {Object} income
 */
function wwPostValidatedIncome(income) {

  wwValidateIncome(income);

  wwPostIncome(income);

}

/**
 * Returns the number of income records currently due.
 *
 * @returns {number}
 */
function wwGetDueIncomeCount() {

  return wwGetDueIncome().length;

}

/**
 * Returns true if at least one income record is due.
 *
 * @returns {boolean}
 */
function wwHasDueIncome() {

  return wwGetDueIncomeCount() > 0;

}

/**
 * Returns the total amount of all income currently due.
 *
 * @returns {number}
 */
function wwGetDueIncomeTotal() {

  return wwGetDueIncome().reduce(function(total, income) {
    return total + (Number(income.amount) || 0);
  }, 0);

}

/**
 * Returns a summary object for the Income sheet.
 *
 * @returns {Object}
 */
function wwGetIncomeSummary() {

  return {
    totalIncomeRecords: wwGetIncomeCount(),
    activeIncomeRecords: wwGetActiveIncome().length,
    dueIncomeRecords: wwGetDueIncomeCount(),
    dueIncomeAmount: wwGetDueIncomeTotal()
  };

}
