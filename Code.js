// Function to process the form submission
function processExpenseForm(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getLastRow();
  const expenseData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const timestamp = expenseData[0];
  const employeeEmail = expenseData[1];
  const employeeName = expenseData[2];
  const date = expenseData[3];
  const category = expenseData[4];
  const amount = parseFloat(expenseData[5]);
  const description = expenseData[6];
  const receiptUrl = expenseData[7];

  const categories = {
    'Travel': 'Travel Expenses',
    'Office Supplies': 'Office Expenses',
    'Meals': 'Food Expenses',
    'Other': 'Miscellaneous'
  };

  const categorizedAs = categories[category] || 'Uncategorized';
  sheet.getRange(row, 9).setValue(categorizedAs);

  updateBudgetTracking(categorizedAs, amount);

  const managerEmail = 'meteorsee1108@gmail.com';  // Manager's email
  const subject = 'New Expense Report Submitted';
  const webAppUrl = ScriptApp.getService().getUrl();

  const approvalUrl = `${webAppUrl}?row=${row}&action=approve`;
  const rejectionUrl = `${webAppUrl}?row=${row}&action=reject`;

  const body = `
    A new expense report has been submitted by ${employeeName}.

    Date: ${date}
    Category: ${category} (Categorized as: ${categorizedAs})
    Amount: $${amount.toFixed(2)}
    Description: ${description}
    Receipt: <a href="${receiptUrl}">View Receipt</a>

    Please review and approve or reject the expense:
    <a href="${approvalUrl}">Approve</a> | <a href="${rejectionUrl}">Reject</a>
  `;

  MailApp.sendEmail(managerEmail, subject, '', {htmlBody: body});
}

// Function to update the budget tracking
function updateBudgetTracking(category, amount) {
  const budgetSheetName = 'Budget Tracking';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let budgetSheet = ss.getSheetByName(budgetSheetName);

  // Create the budget sheet if it doesn't exist
  if (!budgetSheet) {
    budgetSheet = ss.insertSheet(budgetSheetName);
    budgetSheet.appendRow(['Category', 'Amount Spent']);
  }

  const budgetData = budgetSheet.getDataRange().getValues();
  let categoryFound = false;

  for (let i = 1; i < budgetData.length; i++) {
    if (budgetData[i][0] === category) {
      let currentAmount = parseFloat(budgetData[i][1]) || 0;
      budgetSheet.getRange(i + 1, 2).setValue(currentAmount + amount);
      categoryFound = true;
      break;
    }
  }

  if (!categoryFound) {
    budgetSheet.appendRow([category, amount]);
  }
}

// Function to handle the approval/rejection
function doGet(e) {
  const params = e.parameter;
  const row = parseInt(params.row);
  const action = params.action;

  const htmlOutput = HtmlService.createTemplateFromFile('responseForm');
  htmlOutput.row = row;
  htmlOutput.action = action;

  return htmlOutput.evaluate()
    .setTitle(action === 'approve' ? 'Approve Expense' : 'Reject Expense')
    .setWidth(400)
    .setHeight(300);
}

// Function to process the approval/rejection form submission
function doPost(e) {
  const params = e.parameter;
  const row = parseInt(params.row);
  const action = params.action;
  const comments = params.comments || '';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const employeeEmail = sheet.getRange(row, 2).getValue();
  const employeeName = sheet.getRange(row, 3).getValue();
  const approvalStatus = (action === 'approve') ? 'Approved' : 'Rejected';

  sheet.getRange(row, 10).setValue(approvalStatus);
  sheet.getRange(row, 11).setValue(comments);

  const subject = `Expense Report ${approvalStatus}`;
  const body = `Hi ${employeeName},

Your expense report has been ${approvalStatus.toLowerCase()}.

Manager's comments: ${comments}

Best regards,
Expense Management Team`;

  MailApp.sendEmail(employeeEmail, subject, body);

  return HtmlService.createHtmlOutput(`Expense report has been ${approvalStatus.toLowerCase()}.`);
}

// Trigger to process the form submission
function createTrigger() {
  const formId = '1C8F-BRbeE23WJGYogxHcvhev6YPmcASFO2WVSSyNBbs';
  const form = FormApp.openById(formId);

  // Remove existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  for (let trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processExpenseForm') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create a new trigger
  ScriptApp.newTrigger('processExpenseForm')
    .forForm(form)
    .onFormSubmit()
    .create();
}

// Run this function once to create the trigger
function setup() {
  createTrigger();
}

// Function to set up the budget tracking sheet (run once)
function setupBudgetTracking() {
  const budgetSheetName = 'Budget Tracking';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let budgetSheet = ss.getSheetByName(budgetSheetName);
  if (!budgetSheet) {
    budgetSheet = ss.insertSheet(budgetSheetName);
    budgetSheet.appendRow(['Category', 'Amount Spent']);
  }
}
