# CertiFlow Workspace - Enterprise Certificate Automation Suite

CertiFlow Workspace is a highly responsive, secure, and sovereign SaaS-style automation software built entirely on top of the Google Workspace Infrastructure (Google Apps Script, Sheets, Slides, Drive, and Gmail API). 

It allows institutions to seamlessly parse recipient data from both remote Google Sheets and local Excel (.xlsx) binaries to generate verified certificates complete with unique tracking QR codes.

## Key Technical Architectures
- **Dual Data Stream Pipeline:** Parses remote cloud sheets (updating tracking statuses dynamically) as well as binary Excel tables locally on-the-fly.
- **Sovereign Execution Layer:** Zero third-party backend servers; operations run securely under the user's specific Google Account token authorizations.
- **Isolated Slide Indexing Engine:** Intelligently isolates and trims unwanted template nodes dynamically to compile lightweight single-page certificate outputs.
- **Aesthetic Frontend Guardrails:** Built with custom dynamic form state switches, comprehensive UI validation, and responsive error border structures using Tailwind CSS.
- **Time-Driven Scheduling Cron:** Fully decoupled cron module capable of queuing pipeline releases for future execution dates.

## Technology Stack
- **Frontend:** HTML5, JavaScript (ES6+), Tailwind CSS Architecture
- **Backend Infrastructure:** Google Apps Script Runtime Engine
- **Core Integrations:** DocumentApp, SlidesApp, DriveApp, GmailApp, Advanced Drive API v3

## Production Installation & Setup
1. Create an empty folder in your Google Drive to store the generated PDFs and copy its Folder ID.
2. Open your Google Sheet, navigate to **Extensions -> Apps Script**, and paste the code from `Code.gs`.
3. Create a new HTML file named `Index.html` in the Apps Script editor and paste the frontend code.
4. Replace the `OUTPUT_FOLDER_ID` placeholder at the top of `Code.gs` with your real Google Drive folder ID.
5. Enable **Drive API v3** under the Advanced Services block in your script editor.
6. Deploy the software ecosystem natively as a Web App targeted with "Anyone with Google Account" access parameters.
