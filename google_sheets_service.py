import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

class GoogleSheetsService:
    def __init__(self):
        self.service = None
        self._initialize_service()
    
    def _initialize_service(self):
        """Initialize the Google Sheets API service"""
        try:
            # Check if we have a service account key file
            if os.path.exists('service_account_key.json'):
                credentials = service_account.Credentials.from_service_account_file(
                    'service_account_key.json',
                    scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
                )
            else:
                # Try to get credentials from environment variable
                service_account_info = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
                if service_account_info:
                    credentials = service_account.Credentials.from_service_account_info(
                        json.loads(service_account_info),
                        scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
                    )
                else:
                    raise ValueError("No Google service account credentials found")
            
            self.service = build('sheets', 'v4', credentials=credentials)
            print("Google Sheets service initialized successfully")
            
        except Exception as e:
            print(f"Error initializing Google Sheets service: {e}")
            raise
    
    def read_sheet(self, spreadsheet_id, range_name):
        """
        Read data from a Google Sheet
        
        Args:
            spreadsheet_id (str): The ID of the spreadsheet
            range_name (str): The A1 notation range to read
            
        Returns:
            list: List of rows, where each row is a list of cell values
        """
        try:
            sheet = self.service.spreadsheets()
            result = sheet.values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            return values
            
        except HttpError as error:
            print(f"An error occurred: {error}")
            raise
    
    def get_sheet_info(self, spreadsheet_id):
        """
        Get information about a Google Sheet
        
        Args:
            spreadsheet_id (str): The ID of the spreadsheet
            
        Returns:
            dict: Spreadsheet metadata
        """
        try:
            sheet = self.service.spreadsheets()
            result = sheet.get(spreadsheetId=spreadsheet_id).execute()
            
            return result
            
        except HttpError as error:
            print(f"An error occurred: {error}")
            raise
    
    def get_sheet_names(self, spreadsheet_id):
        """
        Get the names of all sheets in a spreadsheet
        
        Args:
            spreadsheet_id (str): The ID of the spreadsheet
            
        Returns:
            list: List of sheet names
        """
        try:
            sheet_info = self.get_sheet_info(spreadsheet_id)
            sheets = sheet_info.get('sheets', [])
            
            return [sheet['properties']['title'] for sheet in sheets]
            
        except HttpError as error:
            print(f"An error occurred: {error}")
            raise
    
    def batch_read_sheets(self, spreadsheet_id, ranges):
        """
        Read multiple ranges from a Google Sheet in a single request
        
        Args:
            spreadsheet_id (str): The ID of the spreadsheet
            ranges (list): List of A1 notation ranges to read
            
        Returns:
            dict: Dictionary with ranges as keys and data as values
        """
        try:
            sheet = self.service.spreadsheets()
            result = sheet.values().batchGet(
                spreadsheetId=spreadsheet_id,
                ranges=ranges
            ).execute()
            
            value_ranges = result.get('valueRanges', [])
            
            # Create a dictionary mapping ranges to their data
            data = {}
            for i, range_name in enumerate(ranges):
                if i < len(value_ranges):
                    data[range_name] = value_ranges[i].get('values', [])
                else:
                    data[range_name] = []
            
            return data
            
        except HttpError as error:
            print(f"An error occurred: {error}")
            raise 