┌────────────────────────┐
│   Customer / Client    │
└───────────┬────────────┘
            │ 1. Submits complaint (Web Form)
            ▼
┌────────────────────────┐
│  Google Apps Script    │ ── 2. Generates Unique Ticket ID (LockService: BHT-xxxx)
│   Backend (Code.gs)    │ ── 3. Appends raw record to Sheet ("Complaints")
└───────────┬────────────┘ ── 4. Sends confirmation email (GmailApp / MailApp)
            │
            ├─────────────────────────────────────────┐
            │                                         │
            ▼ (Assigns Ticket)                        ▼ (Submits Report)
┌────────────────────────┐               ┌────────────────────────┐
│  Coordinator Profile   │               │ Field Engineer Profile │
└───────────┬────────────┘               └───────────┬────────────┘
            │                                         │
            │ Selects Primary & Co-Engineer           │ Uploads photos (Drive API)
            │ (Migrates legacy names via script)      │ Inputs charges (KM, food, extra)
            ▼                                         ▼
┌─────────────────────────────────────────────────────────────────┐
│                     Google Sheets Database                      │
│   [Complaints]   [ServiceReports]   [Invoices]   [UploadedFiles]│
└───────────────────────────┬─────────────────────────────────────┘
                            │
                            │ CacheSync (CacheService chunking)
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│               Admin / Accounts / HR Dashboard                   │
│         - Tracks billing, invoice uploads, & reports            │
└─────────────────────────────────────────────────────────────────┘


## 🔄 System Workflow & Lifecycle

1. **Complaint Registration:**
   - Customer or Coordinator fills out the service request form via the Web App.
   - The system acquires a script lock (`LockService`) to sequentially generate a unique Ticket ID (e.g., `BHT-0001`).
   - The ticket data is written to the `Complaints` tab in Google Sheets, and an automated email receipt is sent.

2. **Engineer Assignment & Migration:**
   - The Coordinator assigns a primary engineer (e.g., `VIVEK`) and an optional accompanying person.
   - Built-in live sync (`migrateLegacyNamesInSheets`) automatically normalizes historical technician names in legacy records.

3. **Field Service & File Upload:**
   - The engineer performs the on-site service and submits a **Service Report**.
   - Equipment photos and documentation are converted to Base64 in the frontend and uploaded to dedicated Google Drive folders (`BHT Uploads`).
   - File metadata (URL, file ID, mime type) is cataloged in the `UploadedFiles` sheet tab.

4. **Caching & Account Reconciliation:**
   - Read requests utilize custom chunked caching (`CacheService`) to drastically improve dashboard load times.
   - The Accounts team views completed service reports, generates invoice records, and links payment status.
