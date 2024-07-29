# Expense Management System

This project is an Expense Management System built using Google Apps Script. It allows employees to submit expense reports via Google Forms, automates the categorization and tracking of expenses, and includes an approval workflow where managers can approve or reject expenses with comments.

## Features

- **Expense Submission via Google Forms:** Employees can submit expense reports using Google Forms, which integrate seamlessly with Google Sheets.
- **Automated Expense Categorization:** Automatically categorize expenses based on predefined rules and store them in a central Google Sheet.
- **Real-Time Budget Tracking:** Provide real-time updates on spending against budgets with visual dashboards in Google Sheets.
- **Expense Approval Workflow:** Automate the approval process, notifying managers via email and updating approval status in Google Sheets.

## Setup

### Google Sheets and Forms

1. **Google Form:**
   - Create a new Google Form.
   - Add questions corresponding to the columns in your Google Sheet: `Email Address`, `Employee Name`, `Date`, `Expense Category`, `Amount`, `Description`, `Upload receipt`.

2. **Link Form to Sheet:**
   - Click on the "Link to Sheet" button so that form responses are recorded in the Google Sheet.

3. **Google Sheet:**
   - Add the following columns to the first sheet: `Categorized As`, `Approval Status`, `Manager Comments`.

### Google Apps Script

1. **Script Setup:**
   - Open the Google Sheet.
   - Go to `Extensions` -> `Apps Script`.
   - Delete any existing code and paste the provided code from this repository into the script editor.
   - Save the script.

2. **Deploy the Web App:**
   - Go to `Deploy` -> `New deployment`.
   - Select `Web app` and follow the instructions to deploy.
   - Copy the web app URL for use in the email notifications.

3. **Budget Tracking Sheet:**
   - Run the `setupBudgetTracking` function to create the `Budget Tracking` sheet if it doesn't already exist.

4. **Create Trigger:**
   - Run the `setup` function to create the necessary triggers for processing form submissions.

### HTML Script

1. **Create HTML File:**
   - In the Google Apps Script project, click on the + icon next to "Files" and select "HTML".
   - Name the new file `responseForm.html`.
   - Delete any existing code and paste the provided HTML code from this repository into the script editor.
   - Save the script.
