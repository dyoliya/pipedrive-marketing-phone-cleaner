# 📊 Pipedrive Marketing Phone Cleaner

> **Cleans and normalizes phone numbers from Pipedrive deals, removes duplicates and opt-outs, and outputs a clean list for marketing use.**

---

![Version](https://img.shields.io/badge/version-1.5.1-ffab4c?style=for-the-badge&logo=python&logoColor=white)
![Python](https://img.shields.io/badge/python-3.9%2B-273946?style=for-the-badge&logo=python&logoColor=ffab4c)
![Status](https://img.shields.io/badge/status-active-273946?style=for-the-badge&logo=github&logoColor=ffab4c)

---

## ✨ Features

- **Automatic Folder Management:** Creates and uses `for_processing` as the input folder for input files and `output` as the destination folder for cleaned results.
- **Excel and CSV Input Support:** Accepts both `.xlsx` and `.csv` files, with automatic UTF-8/CP1252 fallback for CSV encoding.
- **Column Name Flexibility:** Normalizes and aliases column headers so files exported directly from Pipedrive and files pulled via the Pipedrive API are both recognized (e.g., `Phone` → `Person - Phone - Work`, `ID` → `Deal - ID`).
- **Multi-Source Phone Number Handling:** Processes multiple phone fields from Pipedrive deal data including work, home, mobile, numbered, and archived numbers.
- **Normalization and Validation:** Normalizes phone numbers by stripping non-digit characters and standardizing to 10-digit US-style numbers (removes leading ‘1’ if present).
- **Opt-Out List Integration:** Downloads and checks phone numbers against opt-out lists stored in Google Drive (`PD Remove DNC.xlsx`, `Sly DNC.xlsx`, `C3.xlsx`, plus `CallTextOut-7d (PD).xlsx` or `CallOut-14d+TextOut-30d (Cold).xlsx` depending on deal stage), marking flagged numbers with remarks.
- **Existing Pipedrive Phone Cross-Check:** Fetches phone numbers from the Google Drive `pd_phone` folder and flags numbers that exist in other deals with different stages, preventing duplication or conflicting marketing outreach.
- **Duplicate Detection Within Input Files:** Tracks phone numbers processed within the current batch to avoid duplicates across deals, annotating duplicates with appropriate remarks.
- **Detailed Remarks and Reporting:** Provides comprehensive remarks per record, noting phone format issues, opt-out presence, existing deal conflicts, and duplicate status to aid downstream decisions.
- **Robust Error Handling:** Skips problematic rows with clear console warnings.
- **Combined Excel Output per Run:** Consolidates all cleaned results from each run into a single Excel workbook, with each input file saved as its own sheet.
- **Carrier Sheet & Lookup Formula:** Adds an empty `carrier` sheet at the end and applies the formula `=VLOOKUP(C2,carrier!A:C,3,FALSE)` to the **Carrier** column in each sheet (column **C** refers to the **Phone Number** column).
- **Timestamped Filenames:** Output file names now follow this format: `yyyymmdd_HHMMSS_pd_mktg_combined_output.xlsx` for clear version tracking.

---
## 🧠 Logic Flow
1. User provides files
   - User places Excel (`.xlsx`) or CSV (`.csv`) files into the for_processing folder.
   - CSV files are read as UTF-8, falling back to CP1252 if decoding fails.
2. Tool normalizes and checks file requirements
   - Column headers are trimmed of stray whitespace and aliased (e.g., `Phone` → `Person - Phone - Work`, `ID` → `Deal - ID`, `Contact person` → `Deal - Contact person`) so both Pipedrive exports and API pulls are recognized.
   - Each file must include the following columns (after aliasing):
     - Deal - ID
     - Deal - Contact person
     - Deal - Owner
     - Deal - County
     - Deal - Stage
   - Each file must also contain at least one supported phone column (for example: Person - Phone - Mobile, Person - Phone - Work, etc.).
   - If required columns are missing, the file is skipped.
3. Tool scans supported phone fields
   - Phone values are read from the supported phone columns.
   - Multiple phone numbers within a cell are allowed if separated by commas; blank entries caused by extra commas are ignored.
4. Phone numbers are normalized
   - Symbols and spaces are removed so only digits remain.
   - If a number becomes 11 digits starting with 1, the leading 1 is removed.
   - Only 10-digit numbers are treated as valid after normalization.
   - Invalid formats are recorded in the Remarks column.
5. Opt-out checking occurs first (Google Drive sources)
   - The tool downloads opt-out lists from Google Drive and compares each valid phone number.
   - The specific lists used depend on the deal stage:
     - If Deal - Stage is one of: Cold Deals - Priority 2, Update Offer - TM Recycling, Document Auditing - Recycling, Generate Offer - Recycling, Large Recycled Lead
       - PD Remove DNC.xlsx
       - Sly DNC.xlsx
       - C3.xlsx
       - CallOut-14d+TextOut-30d (Cold).xlsx
     - All other stages
       - PD Remove DNC.xlsx
       - Sly DNC.xlsx
       - C3.xlsx
       - CallTextOut-7d (PD).xlsx
   - Any phone found in these lists is recorded in Remarks.
6. Existing Pipedrive phone check
   - Existing phone records are loaded from the Google Drive pd_phone folder.
   - If a phone already exists under another Deal ID in a different deal stage, the number is treated as not allowed and recorded in Remarks.
7. Duplicate check within the current run
   - All valid phone numbers processed during the run are tracked across all files.
   - If a phone appears under another Deal ID in the same run, it is marked as a duplicate in Remarks.
   - The first Deal ID that used the phone is treated as the reference record for that number.
8. Phone selection per row
   - Only one phone number is retained per row.
   - The retained number is the first valid and unique phone encountered based on the order of phone fields.
   - If critical issues exist (opt-out, PD conflict, or duplicate), the phone number is removed and the issue remains documented in Remarks.
9. Output generation
   - A single timestamped Excel report is created in the output folder.
   - Each processed input file appears as a separate sheet within the combined report.
   - A carrier sheet is included for optional lookup.
10. Carrier lookup behavior
      - The Carrier column contains a VLOOKUP formula referencing the carrier sheet.
      - Carrier values populate only when the carrier sheet is filled with lookup data.
---

## 🚀 Installation and Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/pipedrive-marketing-cleaning-tool.git
   cd pipedrive-marketing-cleaning-tool

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt

3. **Folder Structure**
<pre>project/
│
├── for_processing/                # Input Excel/CSV files
├── output/                        # Cleaned results
├── config/                        # Configuration files
│   ├── gdrive_client.py           # Google Drive connection logic
│   ├── gdrive_credentials.json    # Google OAuth client credentials
│   ├── gdrive_files.json          # Map of opt-out list names to Drive file IDs
│   ├── gdrive_folders.json        # Map of Drive folder names to folder IDs
│   └── token.json                 # Cached OAuth token (auto-generated)
├── tool_ui.py                      # GUI interface
├── pd_marketing_cleaning_tool.py   # Main script
└── requirements.txt                # Dependencies
</pre>

Before running the tool, you need to set up Google Drive API credentials so the app can securely download the opt-out lists and existing phone records.

4. **Set Up Configuration**

    4.1. **Create a Google Cloud project and OAuth credentials**

   - Go to the [Google Cloud Console](https://console.cloud.google.com/) and create (or select) a project.
   - Enable the **Google Drive API** for that project.
   - Under **APIs & Services > Credentials**, create an **OAuth client ID** of type **Desktop app**.
   - Download the generated credentials file and save it as `config/gdrive_credentials.json`.

   4.2. **Configure file and folder IDs**

   - In Google Drive, locate the opt-out list files (e.g., `PD Remove DNC.xlsx`, `Sly DNC.xlsx`, `C3.xlsx`, `CallTextOut-7d (PD).xlsx`, `CallOut-14d+TextOut-30d (Cold).xlsx`) and the `pd_phone` folder, and note each item's Drive file/folder ID (found in its sharing URL).
   - Add the file name → file ID mappings to `config/gdrive_files.json`.
   - Add the folder name → folder ID mapping (key `pd_phone`) to `config/gdrive_folders.json`.

   4.3. **Authorize the app**

   - The first time the tool runs, it will open a browser window prompting you to sign in with the Google account that has access to the Drive files above.
   - After you approve access, the tool saves the resulting token to `config/token.json` and reuses it (refreshing automatically) on future runs.


5. **Compile the tool**
   ```bash
   pyinstaller --onefile --windowed --add-data "config/gdrive_credentials.json;config" tool_ui.py
---

## 🖥️ User Guide
1. **First-Time Setup**

  * When you open the tool for the first time, it will automatically create two folders:
      * for_processing – where you place files to be cleaned
      * output – where cleaned files are saved

> :bulb: Tip: Place the .exe file in the location where you want these folders to be stored before running it the first time.

2. **Opening the Tool**

  * Double-click the program file to start it.


3. **Checking Your Files**

  * When the tool opens, it will show a list of files currently in the for_processing folder.
  * If the list is wrong or empty:
      * Click Open Folder to open for_processing and adjust your files.
      * Then click Refresh in the tool to reload the list.
  * Make sure the input files and Google Drive files are up to date before running the tool.

> :pencil: Note:
> * Only the files inside the input folder will be processed. 
> * Excel (.xlsx) and CSV (.csv) files are supported. 



4. **Running the Cleaning**

  * Make sure the list shows the correct files.
  * Click RUN TOOL.
  * A “Processing” window will appear — don’t close it.
  * Wait until you see “Processing finished successfully!”


5. **Getting the Results**

  * When processing is done, you will be asked if you want to open the output folder.
  * Click Yes to see your cleaned files.
  * The cleaned files will always be saved in the output folder.


> :warning: **Important Notes**
>
> * Do not close the “Processing” popup before it finishes — this can interrupt the process.
> * You cannot run the tool twice at the same time.
> * The tool will ignore any file that is not an Excel (.xlsx) or CSV (.csv) file. 

---

## 👩‍💻 Credits
- **2025-08-06**: Project created by **Julia** ([@dyoliya](https://github.com/dyoliya))  
- 2025–present: Maintained by **Julia** for **Community Minerals II, LLC**
