import win32com.client as win32
import pandas as pd
import os
import logging
from datetime import datetime



class BillingFilesDataExtractor:
    def __init__(self, base_path=None, required_directory="Absolute Billing and Payroll", password="abs$0321$S"):
        # Define the base path without the required directory
        base = base_path or f"C:\\Users\\{os.getlogin()}\\OneDrive - Ability Home Health, LLC"
        self.required_directory = required_directory
        self.password = password

        # Find the required directory within the base path
        self.base_path = self.find_directory(base, self.required_directory, max_depth=5)
        if self.base_path is None:
            logging.error(f"Required directory '{self.required_directory}' not found within base path: {base}")
            raise FileNotFoundError(f"Required directory '{self.required_directory}' not found within base path: {base}")
        else:
            logging.debug(f"Using base path: {self.base_path}")

        # Define files and corresponding sheets to be processed
        self.files_info = {
            "Anthem ATTC&HMK 2023-2024.xlsx": ["Indy ATTC&HMK 2024", "SB ATTC&HMK 2024"],
            "United ATTC&HMK 2023-2024.xlsx": ["Indy ATTC&HMK 2024", "SB ATTC&HMK 2024"],
            "Humana ATTC&HMK 2023-2024.xlsx": ["Indy ATTC&HMK 2024", "SB ATTC&HMK 2024"],
            "MDC ATTC&HMK 2023-2024.xlsx": ["Indy ATTC&HMK 2024", "SB ATTC&HMK 2024"],
            "Units Record CHOICE South Bend 2023-2024.xlsx": ["Units Record ATTC&HMK 2024"],
            "Units Record CHOICE Indianapolis 2023-2024.xlsx": ["Units Record ATTC&HMK 2024"],
            "Units Record IHCC 2023-2024.xlsx": ["Units Record IHCC 2024"],
            "Units Record NS 2023-2024.xlsx": ["Units Record NUTS 2024"],
            "Units Record PERS 2023-2024.xlsx": ["Units Record PERS 2024"],
            "Units Record SFC 2023-2024.xlsx": ["Units Record SFC 2024"],
        }

    def find_file(self, base_path, filename, max_depth=5):
        logging.debug(f"Searching for file '{filename}' in '{base_path}' up to depth {max_depth}")
        
        def scan_directory(path, current_depth):
            if current_depth > max_depth:
                return None
            try:
                with os.scandir(path) as it:
                    for entry in it:
                        if entry.is_file() and entry.name.lower() == filename.lower():
                            logging.debug(f"Found file: {entry.path}")
                            return entry.path
                        elif entry.is_dir():
                            found_file = scan_directory(entry.path, current_depth + 1)
                            if found_file:
                                return found_file
            except PermissionError as e:
                logging.warning(f"PermissionError accessing '{path}': {e}")
                return None
        
        return scan_directory(base_path, 0)

    def find_directory(self, base_path, directory_name, max_depth=5):
        logging.debug(f"Searching for directory '{directory_name}' in '{base_path}' up to depth {max_depth}")
        
        def scan_directory(path, current_depth):
            if current_depth > max_depth:
                return None
            try:
                with os.scandir(path) as it:
                    for entry in it:
                        if entry.is_dir() and entry.name.lower() == directory_name.lower():
                            logging.debug(f"Found directory: {entry.path}")
                            return entry.path
                        elif entry.is_dir():
                            found_dir = scan_directory(entry.path, current_depth + 1)
                            if found_dir:
                                return found_dir
            except PermissionError as e:
                logging.warning(f"PermissionError accessing '{path}': {e}")
                return None
        
        return scan_directory(base_path, 0)

    def convert_pywintypes_to_string(self, value):
        """ Convert pywintypes.datetime to string, otherwise return the original value. """
        if isinstance(value, datetime):
            return value.strftime('%Y-%m-%d')
        return str(value)

    def find_header_row(self, data):
        """ Find the header row in column A that contains 'Medical Record Number' or 'Patient'. 
        If 'Patient' is found, replace it with 'Medical Record Number'. """
        for index, row in enumerate(data):
            if isinstance(row[0], str) and ("Medical Record Number" in row[0] or "Patient" in row[0]):
                # Replace 'Patient' with 'Medical Record Number' in the header row
                row = ["Medical Record Number" if col == "Patient" else col for col in row]
                data[index] = row
                return index
        return None

    def remove_empty_rows(self, data):
        """ Remove rows where column A is blank or None. """
        return [row for row in data if row[0] not in (None, "", " ")]

    def process_excel_data(self, data):
        """ Converts pywintypes.datetime to strings for all values in the extracted data. """
        processed_data = []
        for row in data:
            processed_data.append([self.convert_pywintypes_to_string(value) for value in row])
        return processed_data

    def ensure_unique_headers(self, headers):
        """ Ensure that column headers are unique by appending a suffix if duplicates are found. """
        seen = {}
        for i, col in enumerate(headers):
            if col in seen:
                # Append a suffix to make the column name unique
                headers[i] = f"{col}_{seen[col]}"
                seen[col] += 1
            else:
                seen[col] = 1
        return headers

    def process_billing_files(self):
        """ Extract data from specified sheets, combine all workbook data, and save to separate sheets in one Excel workbook. """
        merged_df = pd.DataFrame()  # Initialize the merged DataFrame

        try:
            logging.debug(f"Base path: {self.base_path}")
            if not os.path.exists(self.base_path):
                logging.error(f"Base path does not exist: {self.base_path}")
                print(f"Base path does not exist: {self.base_path}")
                return None

            try:
                # Use DispatchEx to create a new Excel instance
                excel = win32.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False  # Disable alerts to avoid prompts
                logging.debug("Excel application started successfully.")
            except Exception as e:
                logging.error(f"Failed to create Excel application: {e}")
                print(f"Failed to create Excel application: {e}")
                return None

            try:
                # Open all workbooks in the base path
                for filename, sheets_to_process in self.files_info.items():
                    try:
                        file_path = self.find_file(self.base_path, filename)
                        if file_path is None:
                            logging.warning(f"File not found: {filename}")
                            print(f"File not found: {filename}")
                            continue
                        else:
                            # print(f"File found: {file_path}")
                            logging.info(f"File found: {file_path}")

                        # Open the workbook (read-only mode)
                        wb = excel.Workbooks.Open(file_path, False, True, None, self.password, '', True)  # Read-only
                        available_sheets = [sheet.Name for sheet in wb.Sheets]  # List all available sheets
                        # print(f"Sheets found in {filename}: {available_sheets}")
                        logging.info(f"Processing file: {filename}")

                        # Process only the specified sheets
                        for sheet_name in sheets_to_process:
                            if sheet_name in available_sheets:
                                ws = wb.Sheets(sheet_name)
                                raw_data = ws.UsedRange.Value  # Raw data from the sheet
                                processed_data = self.process_excel_data(raw_data)

                                # Remove rows where column A is blank or None
                                processed_data = self.remove_empty_rows(processed_data)

                                # Find the header row
                                header_row_index = self.find_header_row(processed_data)
                                if header_row_index is not None:
                                    headers = processed_data[header_row_index]  # Use the found header row

                                    # Ensure unique headers to avoid the reindexing error
                                    headers = self.ensure_unique_headers(headers)

                                    data_start_index = header_row_index + 1  # Data starts after header row
                                    data = processed_data[data_start_index:]
                                    scheduling_df = pd.DataFrame(data, columns=headers)

                                    # Add the Source File column at the beginning
                                    scheduling_df.insert(0, 'Source File', filename)

                                    # Log if processing any empty DataFrame
                                    if scheduling_df.empty:
                                        logging.warning(f"Empty DataFrame for sheet {sheet_name} in {filename}")
                                        print(f"Empty DataFrame for sheet {sheet_name} in {filename}")
                                        continue

                                    # Append the DataFrame to the merged_df
                                    merged_df = pd.concat([merged_df, scheduling_df], ignore_index=True)

                                else:
                                    logging.warning(f"Header row not found in sheet {sheet_name} in {filename}")
                                    print(f"Header row not found in sheet {sheet_name} in {filename}")

                            else:
                                logging.warning(f"Sheet {sheet_name} not found in {filename}")
                                print(f"Sheet {sheet_name} not found in {filename}")

                    except Exception as e:
                        logging.error(f"Error processing file {filename}: {e}")
                        print(f"Error processing file {filename}: {e}")
                        continue

                    finally:
                        # Close the workbook properly
                        wb.Close(False)
                        logging.info(f"Closed workbook: {filename}")

            finally:
                # Close the Excel application properly
                excel.Quit()
                del excel
                logging.debug("Excel application closed.")

            # Export the merged DataFrame to a single Excel workbook
            if not merged_df.empty:
                return merged_df
            else:
                return None

        except Exception as e:
            logging.error(f"Error in process_billing_files: {e}")
            print(f"Error: {e}")

# Run the extraction process
if __name__ == "__main__":
    try:
        extractor = BillingFilesDataExtractor()
        df = extractor.process_billing_files()
        if df is not None:
            # Output DataFrame in CSV format without index
            print(df.to_csv(index=False))
        else:
            print("No eligible patient data was extracted.")
    except FileNotFoundError as fnf_error:
        print(fnf_error)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
