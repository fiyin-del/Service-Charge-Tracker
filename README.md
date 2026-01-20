# Service Charge Tracker (Google Sheets Automation)
## Quick start
1. Download templates/Service_Charge_Tracker_Template.xlsx
2. Upload it to Google Sheets
3. Open Extensions → Apps Script
4. Paste the contents of Code.gs
5. Authorise the script

Set the daily email trigger
## Overview
This project is a Google Sheets–based service charge management system built using **Google Apps Script**. It automates payment tracking, coverage calculations, arrears monitoring, and email reminders for upcoming payments.
The solution is designed for organisations managing recurring service charges and needing a lightweight, auditable, and low-cost system without a full backend.
## Key Features 

* **Data Entry Form**: Dedicated sheet for entering resident and payment details
* **Automated Calculations**:
  * Daily rate calculation
  * Coverage period (start/end dates)
  * Next payment due date
  * Running balance and arrears
* **Email Reminders**:
  * Automatic email reminders sent 1 day before payment is due
  * Duplicate reminder prevention
* **Validation & Safeguards**:
  * Prevents submission when required fields are empty
  * Prevents duplicate email notifications
## Sheets Structure
### 1. Data Entry Sheet
Used by staff to submit new payments.
Required fields:
* Resident Name
* Weekly Rate
* Payment Date
* Amount Paid
Optional fields:
* Resident ID
* Email
* Phone
* Notes
A custom menu button triggers the submission script.

### 2. Service Charge Tracker Sheet
Stores all payment records and calculated fields.

Key columns include:

* Weekly Rate
* Payment Date
* Amount Paid
* Days Covered
* Coverage Start / End
* Balance Before & After Payment
* Next Payment Due
* Arrears
* Notes
## Automation Logic

### Payment Submission
* Appends data from the Data Entry sheet to the tracker
* Automatically runs coverage and balance calculations
* Clears the entry form only after successful submission

### Coverage Calculation
* Converts weekly rates into daily rates
* Calculates how many days a payment covers
* Determines next payment due date
* Carries balances forward row by row

### Email Reminders
* Runs on a time-based trigger
* Sends reminders when payment is due the next day
* Normalises dates to avoid US/UK format issues
* Logs sent reminders to prevent duplicates
  
## Technologies Used
* Google Sheets
* Google Apps Script (JavaScript)
* Gmail (MailApp)

## Setup Instructions
1. Create a Google Sheet with the required sheet names
2. Open **Extensions → Apps Script**
3. Paste the provided script into the editor
4. Save and authorise the project
5. Set a **time-based trigger** for `sendPaymentReminders`
6. Use the custom menu to submit payments

## Intended Use
This project is suitable for:
* Housing providers
* Charities
* Supported accommodation services
* Administrative or finance teams

## Limitations
* Requires Google Workspace
* Email reminders depend on triggers
  
## Disclaimer
This tool is intended for internal administrative use and does not replace professional accounting software.

## Author
Mofiyinfoluwa Olufala
