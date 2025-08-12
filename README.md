# 📊 Pipedrive Marketing Cleaning Tool

> **From the list of Pipedrive deals for marketing (with phone numbers), this tool normalizes the Excel data, checks against opt-outs, verifies against existing deals with different stages, and deduplicates to ensure unique valid phone numbers for marketing purposes.**

---

![Version](https://img.shields.io/badge/version-1.0.1-ffab4c?style=for-the-badge&logo=python&logoColor=white)
![Python](https://img.shields.io/badge/python-3.9%2B-273946?style=for-the-badge&logo=python&logoColor=ffab4c)
![Status](https://img.shields.io/badge/status-active-273946?style=for-the-badge&logo=github&logoColor=ffab4c)

---

## ✨ Features

- **Automatic Folder Management:** Creates and uses `for_processing` as the input folder for Excel files and `output` as the destination folder for cleaned results.
- **Multi-Source Phone Number Handling:** Processes multiple phone fields from Pipedrive deal data including work, home, mobile, and archived numbers.
- **Normalization and Validation:** Normalizes phone numbers by stripping non-digit characters and standardizing to 10-digit US-style numbers (removes leading ‘1’ if present).
- **Opt-Out List Integration:** Downloads and checks phone numbers against opt-out lists stored in Dropbox (e.g., DNC (Cold-PD).csv and CallTextOut-7d (PD).csv), marking flagged numbers with remarks.
- **Existing Pipedrive Phone Cross-Check:** Fetches phone numbers from Dropbox’s pd_phone folder and flags numbers that exist in other deals with different stages, preventing duplication or conflicting marketing outreach.
- **Duplicate Detection Within Input Files:** Tracks phone numbers processed within the current batch to avoid duplicates across deals, annotating duplicates with appropriate remarks.
- **Customizable Deal Stage Handling:** Segregates and formats cleaned data differently based on deal stages such as Junior Sales, Sales, and Conversion Qualifying.
- **Detailed Remarks and Reporting:** Provides comprehensive remarks per record, noting phone format issues, opt-out presence, existing deal conflicts, and duplicate status to aid downstream decisions.
- **Robust Error Handling:** Skips problematic rows with clear console warnings.
- **Outputs Cleaned Data:** Saves cleaned and processed data files suffixed with `_cleaned.xlsx` into the output folder for easy access and further marketing workflows.

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
├── for_processing/              # Input Excel files
├── output/                       # Cleaned results
├── config/                       # Configuration files
│   ├── .env                       # Environment variables
│   └── dropbox_config.py          # Dropbox connection logic
├── tools/                       # Configuration files
│   └── dropbox_token_generator.py         # Dropbox refresh token generator
├── tool_ui.py                     # GUI interface
├── pd_marketing_cleaning_tool.py  # Main script
└── requirements.txt               # Dependencies
</pre>

Before running the tool, you need to set up your Dropbox app credentials and generate a **refresh token** to allow the app to access your Dropbox account securely.

4. **Set Up Configuration**

    Before running the tool, you need to set up your Dropbox app credentials and generate a **refresh token** to allow the app to access your Dropbox account securely.

    4.1. **Create a Dropbox App**

   - Go to the [Dropbox App Console](https://www.dropbox.com/developers/apps).
   - Click **Create App**.
   - Select **Scoped access** and **Full Dropbox** or **App folder** depending on your needs.
   - Give your app a unique name.
   - In the app settings, add this as your **Redirect URI**:  
     `http://localhost:8080`  
     *(Make sure it matches the `REDIRECT_URI` in your token generation script.)*
   - Save the app.

   4.2. **Get Your App Credentials**

   - Copy your **App Key** and **App Secret** from the app settings.
   - Paste these values into your `.env` file like so:

     ```env
     DROPBOX_CONVERSION_APP_KEY=your_app_key_here
     DROPBOX_CONVERSION_APP_SECRET=your_app_secret_here
     ```

   4.3. **Generate the Refresh Token**

   - Use the included script (`generate_refresh_token.py`) to perform the OAuth flow and obtain a refresh token.
   - Edit the script and fill in your app credentials at the top:
      ```python
     APP_KEY = "your_app_key_here"       # Dropbox App Key
     APP_SECRET = "your_app_secret_here" # Dropbox App Secret
     REDIRECT_URI = "http://localhost:8080"
     ```
   - Run the script locally (it will open a browser window and prompt for Dropbox login).
   - After successful login, it will output the **refresh token** in the **dropbox_tokens.json** file.
   - Copy the refresh token and add it to your `.env` file:

     ```env
     DROPBOX_CONVERSION_REFRESH_TOKEN=your_refresh_token_here
     ```

   4.4. **Place the `.env` file inside the `config/` folder**

   - The tool will automatically load these environment variables from `config/.env`.


5. **Compile the tool**
   ```bash
   pyinstaller --onefile --windowed --add-data "config/.env;config" tool_ui.py
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
  * Make sure the files from the dropbox are updated.

> :pencil: Note:
> * Only the files inside the input folder will be processed. 
> * Only Excel files (.xlsx) are supported. 



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
> * The tool will ignore any file that is not an Excel (.xlsx) file. 

---

## 📌 Version History
1.0.1 (Patch)

Improved folder handling (auto-create input folder)

UI Run button now locks during processing

Fixed .env loading for packaged executable

1.0.0

Initial release with core cleaning, deduplication, and opt-out checks.
