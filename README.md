# Telegram Bulk PDF Bot

Private/sample certificate automation.

## Flow

/start
-> upload .xlsx
-> enter range, e.g. 2-20
-> one PDF per row
-> each PDF gets a QR generated from that row's plain-text data
-> PDFs are sent one-by-one to Telegram

## Required sheet headers

Registration_Number
Legal_Name
Trade_Name
Constitution
Address
Date_Liability
Valid_From
Valid_To
Type_Registration
Approving_Authority
Authority_Name
Authority_Designation
Jurisdictional_Office
Date_Issue

Row 1 = headers. Data starts at row 2.

## Railway

Deploy this repository from GitHub and set:

BOT_TOKEN=YOUR_TELEGRAM_BOT_TOKEN

The included Dockerfile installs Chromium for Playwright.
