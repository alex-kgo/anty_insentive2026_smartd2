import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import os
import config
from dotenv import load_dotenv

load_dotenv()

class GoogleSheetManager:
    def __init__(self):
        self.scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        self.creds = None
        self.client = None
        self.sheet = None # The specific worksheet for the month

    def authenticate(self):
        """Authenticate with Google Sheets API."""
        json_path = os.getenv("GOOGLE_JSON_PATH")
        if not json_path or not os.path.exists(json_path):
             raise Exception(f"Google Service Account JSON not found at: {json_path}")
        
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, self.scope)
        self.client = gspread.authorize(self.creds)

    def get_or_create_sheet(self, year_month):
        """
        Get existing worksheet or create a new one for YYYY-MM.
        year_month: str, e.g., "2026-02"
        """
        spreadsheet_key = os.getenv("GOOGLE_SPREADSHEET_KEY")
        # Open by key if available, otherwise by name (requires name in env or config, missing in current env template, assuming key provided)
        # If key is missing, user might want to open by name "Incentive_Data" or similar.
        # Decisions.md doesn't specify the *file* name, only the sheet (tab) name.
        # I'll assume there is a main spreadsheet file key.
        
        if not spreadsheet_key:
             raise Exception("GOOGLE_SPREADSHEET_KEY not found in .env")

        try:
            sh = self.client.open_by_key(spreadsheet_key)
        except gspread.SpreadsheetNotFound:
             raise Exception("Spreadsheet not found. check GOOGLE_SPREADSHEET_KEY.")

        try:
            worksheet = sh.worksheet(year_month)
            print(f"Opened existing worksheet: {year_month}")
        except gspread.WorksheetNotFound:
            print(f"Creating new worksheet: {year_month}")
            worksheet = sh.add_worksheet(title=year_month, rows="100", cols="20")
            # Add Header
            worksheet.append_row(config.SHEET_HEADERS)
        
        self.sheet = worksheet
        return worksheet

    def upsert_data(self, new_data_list):
        """
        Upsert data into the sheet.
        new_data_list: list of dicts
        """
        if not self.sheet:
            raise Exception("No worksheet selected. Call get_or_create_sheet first.")

        # Read all existing data
        existing_records = self.sheet.get_all_records()
        df = pd.DataFrame(existing_records)
        
        # If sheet is empty (only header), df might be empty
        if df.empty:
            df = pd.DataFrame(columns=config.SHEET_HEADERS)

        # Convert simple list of dicts to DataFrame
        new_df = pd.DataFrame(new_data_list)
        
        # Map columns
        rename_map = {
            "date": "날짜",
            "code": "코드",
            "name": "성명",
            "in_sum": "수신 합계",
            "out_sum": "발신 합계",
            "total_sum": "총합계"
        }
        # Use a copy to avoid SettingWithCopy on slice
        new_df = new_df.rename(columns=rename_map).copy()
        
        # Ensure types are strings for key creation
        if not df.empty:
            df['날짜'] = df['날짜'].astype(str)
            df['코드'] = df['코드'].astype(str)
        
        new_df['날짜'] = new_df['날짜'].astype(str)
        new_df['코드'] = new_df['코드'].astype(str)

        # Create Keys
        # Using .loc to avoid warnings
        if not df.empty:
            df.loc[:, 'key'] = df['날짜'] + "_" + df['코드']
        
        new_df.loc[:, 'key'] = new_df['날짜'] + "_" + new_df['코드']

        # Merge Logic
        if df.empty:
            final_df = new_df.copy()
        else:
            # Combine old and new
            combined = pd.concat([df, new_df], ignore_index=True)
            
            # Drop duplicates, keeping last (newest)
            final_df = combined.drop_duplicates(subset=['key'], keep='last').copy()
            
        # Drop key column
        if 'key' in final_df.columns:
            final_df.drop(columns=['key'], inplace=True)

        # Sort
        # Ensure we are working with a copy
        final_df = final_df.sort_values(by=['날짜', '코드']).copy()
        
        # Fill NaNs
        final_df = final_df.fillna("")
        
        # Prepare data for gspread
        # Headers + Data
        # We need to ensure columns are in correct order as per config.SHEET_HEADERS
        # Reorder columns
        final_df = final_df[config.SHEET_HEADERS]
        
        data_to_write = [final_df.columns.tolist()] + final_df.values.tolist()
        
        self.sheet.clear()
        self.sheet.update(data_to_write)
        print(f"Upserted {len(new_data_list)} rows. Total rows now: {len(final_df)}")
