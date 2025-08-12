# -------------------------ABOUT --------------------------

# pyinstaller --onefile --windowed --add-data "config/.env;config" tool_ui.py
# Tool: Pipedrive Marketing Cleaning Tool
# Developer: dyoliya
# Created: 2025-08-06

# © 2025 dyoliya. All rights reserved.

# ---------------------------------------------------------


import os
import re
import dropbox
from glob import glob
from datetime import datetime
import pandas as pd
from tqdm import tqdm
from io import StringIO
from io import BytesIO
from collections import defaultdict
from config.dropbox_config import get_dropbox_client

# ----------------------- DIRECTORIES -----------------------
# input folder
INPUT_FOLDER = "for_processing"
os.makedirs("for_processing", exist_ok=True)
os.makedirs(INPUT_FOLDER, exist_ok=True)

# dropbox folder
DROPBOX_BASE_PATH = "/List Cleaner & JC DNC"
# DROPBOX_BASE_PATH = "/Sales and Conversion Cleaner" # for testing

# output folders
OUTPUT_CLEANED_FOLDER = "output"
os.makedirs("output", exist_ok=True)
os.makedirs(OUTPUT_CLEANED_FOLDER, exist_ok=True)

# ----------------------- PHONE FIELDS -----------------------

PHONE_FIELDS = [
    "Person - Phone - Work",
    "Person - Phone - Home",
    "Person - Phone - Mobile",
    "Person - Phone - Other",
    "Person - Phone 1",
    "Person - Phone 2",
    "Person - Phone 3",
    "Person - Phone 4",
    "Person - Phone 5",
    "Person - Phone 6",
    "Person - Phone 7",
    "Person - Phone 8",
    "Person - Phone 9",
    "Person - Phone 10",
    "Person - Archive - Phone"

]

# ----------------------- DEAL STAGES -----------------------

CONVERSION_QUALIFYING = [
    "Active Leads - Qualifying",
    "Active Leads - Website Email Only",
    "Active Leads - Abandoned",
    "Cold Deals - Priority 2",
    "Cold Deals - Priority 3",
    "Cold Deals - Priority 4"
]

JR_SALES = [
    "Staging",
    "Updated Offer",
    "Contact Attempted - Junior Sales",
    "Waiting on Docs - Junior Sales"
]

SALES = [
    "Staging - Mid Sales",
    "Contact Attempted - Mid Sales",
    "Waiting on Docs - Mid Sales"
]

# ------------------ FUNCTIONS ------------------

def normalize_phone(number):
    return re.sub(r"[^\d]", "", str(number))

def load_opt_out_phone_numbers():
    dbx = get_dropbox_client()
    csv_filenames = ["DNC (Cold-PD).csv", "CallTextOut-7d (PD).csv"]
    numbers = defaultdict(set)  # phone -> set of filenames

    for filename in csv_filenames:
        path = f"{DROPBOX_BASE_PATH}/{filename}"
        try:
            metadata, response = dbx.files_download(path)
            content = response.content.decode("utf-8").strip()
            if not content:
                continue
            df = pd.read_csv(StringIO(content), header=None, dtype=str)
            if df.empty or 0 not in df.columns:
                continue
            nums = df[0].dropna().astype(str).map(normalize_phone)
            for num in nums:
                numbers[num].add(filename)  # add filename to set
        except Exception as e:
            print(f"Error reading Dropbox file {path}: {e}")

    return numbers

def load_pd_phone_numbers():
    dbx = get_dropbox_client()
    pd_phone_folder = f"{DROPBOX_BASE_PATH}/pd_phone"
    pd_phone_numbers = {} 

    try:
        res = dbx.files_list_folder(pd_phone_folder)
        files = res.entries
        while res.has_more:
            res = dbx.files_list_folder_continue(res.cursor)
            files.extend(res.entries)

        for file in files:
            if isinstance(file, dropbox.files.FileMetadata) and file.name.endswith(".xlsx"):
                try:
                    metadata, response = dbx.files_download(file.path_lower)
                    content = response.content
                    df = pd.read_excel(BytesIO(content), engine='openpyxl', dtype=str)
                    df.fillna("", inplace=True)

                    for idx, row in df.iterrows():
                        deal_id = row.get("Deal - ID", "")
                        deal_stage = row.get("Deal - Stage", "")
                        for field in PHONE_FIELDS:
                            raw_phones = str(row.get(field, ""))
                            if not raw_phones.strip():
                                continue
                            for phone in map(str.strip, raw_phones.split(",")):
                                normalized = normalize_phone(phone)
                                if len(normalized) == 11 and normalized.startswith("1"):
                                    normalized = normalized[1:]
                                if len(normalized) == 10 and normalized.isdigit():
                                    # Store deal_id and deal_pipeline for this phone
                                    pd_phone_numbers[normalized] = {"deal_id": deal_id, "deal_stage": deal_stage}

                except Exception as e:
                    print(f"Error reading file {file.path_lower}: {e}")
    except Exception as e:
        print(f"Error reading Dropbox folder {pd_phone_folder}: {e}")

    return pd_phone_numbers

def extract_first_name(contact_person, deal_title):
    name = str(contact_person).strip()

    # If name is missing or "No Name"/"Unknown"
    if not name or re.search(r"\b(no name|unknown)\b", name, re.IGNORECASE):
        title = str(deal_title).strip()
        if not re.match(r"(?i)^no name|^unknown", title):
            return title.split()[0].capitalize() if title else ""
        return ""

    # Always take the first word before space or slash
    first_word = re.split(r"[ /]", name)[0].strip()
    
    return first_word.capitalize()

def extract_deal_owner(deal_owner):
    if pd.isna(deal_owner):
        return ""
    owner = str(deal_owner).strip()
    return owner.split()[0] if owner else ""

def format_deal_county(deal_county):
    if pd.isna(deal_county) or not str(deal_county).strip():
        return ""

    counties = [c.strip() for c in str(deal_county).split(",") if c.strip()]
    grouped = [", ".join(counties[i:i+2]) for i in range(0, len(counties), 2)]

    if len(grouped) == 1:
        return f'"{grouped[0]}"'
    elif len(grouped) == 2:
        return f'"{grouped[0]} and {grouped[1]}"'
    else:
        return f'"{", ".join(grouped[:-1])} and {grouped[-1]}"'


# ------------------ MAIN SCRIPT ------------------

def main():
    seen_normalized_numbers = {}
    dropbox_numbers = load_opt_out_phone_numbers()
    pd_phone_numbers = load_pd_phone_numbers()
    input_files = glob(os.path.join(INPUT_FOLDER, "*.xlsx"))

    for file_path in tqdm(input_files, desc="Processing input files"):
        try:
            df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
            df.fillna("", inplace=True)
            cleaned_rows = []
            

            for idx, row in tqdm(df.iterrows(), total=len(df), desc=f"Processing {os.path.basename(file_path)}", leave=False):
                try:
                    remarks = ""
                    deal_id = row.get("Deal - ID", "")
                    deal_stage = row.get("Deal - Stage", "")
                    contact_person = row.get("Deal - Contact person", "")
                    deal_title = row.get("Deal - Title", "")
                    first_name = extract_first_name(contact_person, deal_title)
                    deal_owner = row.get("Deal - Owner", "")
                    deal_owner_fn = extract_deal_owner(deal_owner)
                    raw_county = row.get("Deal - County", "")
                    formatted_county = format_deal_county(raw_county)

                    duplicates_found = {}
                    phone_to_use = ""
                    disallowed_found = False

                    opt_out_remarks = []
                    pd_phone_remarks = []
                    duplicate_remarks = []
                    format_remarks = []

                    # ------------------ STEP 1: OPT-OUT CHECK FIRST ------------------
                    opt_out_matches = defaultdict(list)
                    remaining_numbers = []

                    for field in PHONE_FIELDS:
                        raw_phones = str(row.get(field, "")).strip()
                        if not raw_phones:
                            continue
                        for phone in map(str.strip, raw_phones.split(",")):
                            normalized = normalize_phone(phone)
                            if len(normalized) == 11 and normalized.startswith("1"):
                                normalized = normalized[1:]
                            if not (len(normalized) == 10 and normalized.isdigit()):
                                format_remarks.append(
                                    f"Phone number {phone} has incorrect format even after normalization"
                                )
                                continue

                            if normalized in dropbox_numbers:
                                for fname in dropbox_numbers[normalized]:
                                    opt_out_matches[fname].append(normalized)
                            else:
                                remaining_numbers.append(normalized)

                    # Add opt-out remarks if any phones found in opt-out
                    if opt_out_matches:
                        # parts = []
                        for fname, nums in opt_out_matches.items():
                            plural = "numbers" if len(nums) > 1 else "number"
                            opt_out_remarks.append(f"Phone {plural} {', '.join(nums)} exist in {fname}")
                        remarks = "; ".join(opt_out_remarks)
                        # Don't set disallowed_found = True yet, because some phones remain


                    # ------------------ STEP 2: PD PHONE CHECK ------------------
                    if not disallowed_found and remaining_numbers:
                        pd_phone_remarks = []
                        

                        for normalized in remaining_numbers:
                            if normalized in pd_phone_numbers:
                                existing_deal = pd_phone_numbers[normalized]["deal_id"]
                                existing_stage = pd_phone_numbers[normalized]["deal_stage"]
                                current_stage = row.get("Deal - Stage", "")
                                if existing_stage != current_stage:
                                    pd_phone_remarks.append(
                                        f"{normalized} exists in Deal ID {existing_deal} on stage {existing_stage} (PD Phone Numbers)"
                                    )

                        if pd_phone_remarks:
                            if remarks:
                                  # If remarks already has something from opt-out step, append
                                remarks += "; " + "; ".join(pd_phone_remarks)
                            else:
                                remarks = "; ".join(pd_phone_remarks)
                            disallowed_found = True


                        # If still no disallow, keep first unique number
                        if not disallowed_found:
                            for normalized in remaining_numbers:
                                if normalized in seen_normalized_numbers:
                                    first_deal = seen_normalized_numbers[normalized]
                                    if first_deal != deal_id:
                                        duplicates_found[normalized] = first_deal
                                        duplicate_remarks.append(
                                            f"Phone number {normalized} already exists in Deal ID {first_deal}"
                                        )
                                else:
                                    seen_normalized_numbers[normalized] = deal_id

                                    if not phone_to_use:
                                        phone_to_use = normalized


                    # ------------------ STEP 3: Remarks & Fallbacks ------------------                    
                    remarks_list = []
                    if format_remarks:
                        remarks_list.append("; ".join(format_remarks))
                    if opt_out_remarks:
                        remarks_list.append("; ".join(opt_out_remarks))
                    if pd_phone_remarks:
                        remarks_list.append("; ".join(pd_phone_remarks))
                    if duplicate_remarks:
                        remarks_list.append("; ".join(duplicate_remarks))

                    remarks = "; ".join(remarks_list)

                    # ------------------ STEP 4: Append to Cleaned Rows ------------------
                    if deal_stage in JR_SALES:
                        cleaned_rows.append({
                            "Carrier": "",
                            "Deal - ID": deal_id,
                            "Phone Number": phone_to_use,
                            "First Name": first_name,
                            "Deal - Value": row.get("Deal - Value", ""),
                            "Deal - Owner": deal_owner_fn,
                            "Deal - County": formatted_county,
                            "Deal - Title": row.get("Deal - Title", ""),
                            "Deal - Stage": deal_stage,
                            "Remarks": remarks
                        })
                    elif deal_stage in SALES:
                        cleaned_rows.append({
                            "Carrier": "",
                            "Deal - ID": deal_id,
                            "Phone Number": phone_to_use,
                            "First Name": first_name,
                            "Deal - Owner": deal_owner_fn,
                            "Deal - County": formatted_county,
                            "Deal - Title": row.get("Deal - Title", ""),
                            "Deal - Stage": deal_stage,
                            "Remarks": remarks
                        })
                    elif deal_stage in CONVERSION_QUALIFYING:
                        cleaned_rows.append({
                            "Carrier": "",
                            "Deal - ID": deal_id,
                            "Phone Number": phone_to_use,
                            "First Name": first_name,
                            "Deal - Stage": deal_stage,
                            "Remarks": remarks
                        })

                except Exception as row_err:
                    print(f"⚠️ Skipping row {idx} in {os.path.basename(file_path)}: {row_err}")


            if cleaned_rows:
                cleaned_df = pd.DataFrame(cleaned_rows)
                output_path = os.path.join(
                    OUTPUT_CLEANED_FOLDER,
                    os.path.basename(file_path).replace(".xlsx", "_cleaned.xlsx")
                )
                cleaned_df.to_excel(output_path, index=False)

        except Exception as e:
            print(f"Error processing {file_path}: {e}")

    
if __name__ == "__main__":
    main()