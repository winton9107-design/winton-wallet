/**
 * ============================================================================
 * Winton Wallet
 * Bills.gs
 * ----------------------------------------------------------------------------
 * Scheduled Bill Management
 * ============================================================================
 */

/**
 * Scheduled Bills column definitions.
 */
const WW_BILLS = Object.freeze({

  COL: Object.freeze({
    ACTIVE: 1,
    NAME: 2,
    ACCOUNT: 3,
    CATEGORY: 4,
    PAYMENT_METHOD: 5,
    AMOUNT: 6,
    DUE_DATE: 7,
    REPEAT: 8,
    SOURCE: 9,
    NOTES: 10,
    LAST_PAID: 11,
    NEXT_DUE: 12
  })

});

/**
 * Returns the Scheduled Bills sheet.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function wwGetBillsSheet() {

  return SpreadsheetApp
    .getActive()
    .getSheetByName(WW.SHEETS.BILLS);

}

/**
 * Returns the last bill row.
 *
 * @returns {number}
 */
function wwGetLastBillRow() {

  const sheet = wwGetBillsSheet();

  return Math.max(sheet.getLastRow(), 2);

}

/**
 * Returns the number of scheduled bills.
 *
 * @returns {number}
 */
function wwGetBillCount() {

  return Math.max(wwGetLastBillRow() - 1, 0);

}

/**
 * Returns true if Scheduled Bills contains data.
 *
 * @returns {boolean}
 */
function wwHasBills() {

  return wwGetBillCount() > 0;

}
/**
 * Returns a bill object from the specified row.
 *
 * @param {number} row
 * @returns {Object|null}
 */
function wwGetBill(row) {

  const sheet = wwGetBillsSheet();

  if (row < 2 || row > sheet.getLastRow()) {
    return null;
  }

  const values = sheet
    .getRange(row, 1, 1, 12)
    .getValues()[0];

  return {
    row: row,
    active: values[WW_BILLS.COL.ACTIVE - 1],
    name: values[WW_BILLS.COL.NAME - 1],
    account: values[WW_BILLS.COL.ACCOUNT - 1],
    category: values[WW_BILLS.COL.CATEGORY - 1],
    paymentMethod: values[WW_BILLS.COL.PAYMENT_METHOD - 1],
    amount: values[WW_BILLS.COL.AMOUNT - 1],
    dueDate: values[WW_BILLS.COL.DUE_DATE - 1],
    repeat: values[WW_BILLS.COL.REPEAT - 1],
    source: values[WW_BILLS.COL.SOURCE - 1],
    notes: values[WW_BILLS.COL.NOTES - 1],
    lastPaid: values[WW_BILLS.COL.LAST_PAID - 1],
    nextDue: values[WW_BILLS.COL.NEXT_DUE - 1]
  };

}

/**
 * Returns every scheduled bill.
 *
 * @returns {Object[]}
 */
function wwGetBills() {

  const bills = [];

  const lastRow = wwGetLastBillRow();

  for (let row = 2; row <= lastRow; row++) {

    const bill = wwGetBill(row);

    if (bill) {
      bills.push(bill);
    }

  }

  return bills;

}

/**
 * Returns only active bills.
 *
 * @returns {Object[]}
 */
function wwGetActiveBills() {

  return wwGetBills().filter(bill => {

    const value = String(bill.active).toLowerCase();

    return value === "true" ||
           value === "yes" ||
           value === "y" ||
           value === "active";

  });

}

/**
 * Returns all bills that are due today or earlier.
 *
 * @returns {Object[]}
 */
function wwGetDueBills() {

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return wwGetActiveBills().filter(bill => {

    if (!(bill.nextDue instanceof Date)) {
      return false;
    }

    const due = new Date(bill.nextDue);
    due.setHours(0, 0, 0, 0);

    return due <= today;

  });

}
/**
 * Posts a scheduled bill to the Register.
 *
 * @param {Object} bill
 */
function wwPostBill(bill) {

  wwPostValidatedTransaction({
    date: new Date(),
    account: bill.account,
    description: bill.name,
    category: bill.category,
    paymentMethod: bill.paymentMethod,
    deposit: "",
    withdrawal: bill.amount,
    source: bill.source,
    notes: bill.notes
  });

  wwUpdateBillAfterPosting(bill);

}

/**
 * Updates a bill after it has been posted.
 *
 * @param {Object} bill
 */
function wwUpdateBillAfterPosting(bill) {

  const sheet = wwGetBillsSheet();
  const today = new Date();

  sheet
    .getRange(bill.row, WW_BILLS.COL.LAST_PAID)
    .setValue(today);

  sheet
    .getRange(bill.row, WW_BILLS.COL.NEXT_DUE)
    .setValue(
      wwCalculateNextDueDate(today, bill.repeat)
    );

}

/**
 * Calculates the next due date based on the bill's
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
 * @returns {Date}
 */
function wwCalculateNextDueDate(date, repeat) {

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
 * Posts every bill that is currently due.
 *
 * @returns {number}
 */
function wwProcessDueBills() {

  const dueBills = wwGetDueBills();

  dueBills.forEach(bill => {

    wwPostBill(bill);

  });

  return dueBills.length;

}
/**
 * Validates a scheduled bill.
 *
 * @param {Object} bill
 */
function wwValidateBill(bill) {

  if (!bill) {
    throw new Error("Bill is required.");
  }

  if (!bill.name) {
    throw new Error("Bill name is required.");
  }

  if (!bill.account) {
    throw new Error("Account is required.");
  }

  if (!bill.amount || Number(bill.amount) <= 0) {
    throw new Error("Bill amount must be greater than zero.");
  }

  if (!(bill.nextDue instanceof Date)) {
    throw new Error("A valid Next Due date is required.");
  }

}

/**
 * Posts a validated scheduled bill.
 *
 * @param {Object} bill
 */
function wwPostValidatedBill(bill) {

  wwValidateBill(bill);

  wwPostBill(bill);

}

/**
 * Returns the number of bills currently due.
 *
 * @returns {number}
 */
function wwGetDueBillCount() {

  return wwGetDueBills().length;

}

/**
 * Returns true if at least one bill is due.
 *
 * @returns {boolean}
 */
function wwHasDueBills() {

  return wwGetDueBillCount() > 0;

}

/**
 * Returns the total amount of all currently due bills.
 *
 * @returns {number}
 */
function wwGetDueBillTotal() {

  return wwGetDueBills().reduce(function(total, bill) {
    return total + (Number(bill.amount) || 0);
  }, 0);

}

/**
 * Returns a summary object for the Scheduled Bills sheet.
 *
 * @returns {Object}
 */
function wwGetBillSummary() {

  return {
    totalBills: wwGetBillCount(),
    activeBills: wwGetActiveBills().length,
    dueBills: wwGetDueBillCount(),
    dueAmount: wwGetDueBillTotal()
  };

}
