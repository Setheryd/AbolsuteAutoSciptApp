import win32com.client as win32
import pandas as pd
import os
import logging
from datetime import datetime
import numpy as np

# Set up logging to a file
logging.basicConfig(
    filename='extract_caregivers.log',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)s:%(message)s'
)

class EmployeeRecordsExtractor:
    def __init__(self, base_path=None, password="abs$1018$B"):
        self.base_path = base_path or f"C:\\Users\\{os.getlogin()}\\OneDrive - Ability Home Health, LLC\\"
        self.password = password
        self.files_info = {
            "Anthem Employee Records 2023-2024.xlsm": "Absolute Operation",
            "Humana Employee Records 2023-2024.xlsm": "Absolute Operation",
            "United Employee Records 2023-2024.xlsm": "Absolute Operation",
            "MDC Employee Records 2023-2024.xlsm": "Absolute Operation"
        }

    def find_file(self, base_path, filename, max_depth=5):
        logging.debug(f"Searching for {filename} in {base_path} up to depth {max_depth}")
        
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
                logging.warning(f"PermissionError: {e}")
                return None
        
        return scan_directory(base_path, 0)

    def convert_pywintypes_to_string(self, value):
        """ Convert pywintypes.datetime to string, otherwise return the original value. """
        if isinstance(value, datetime):
            return value.strftime('%Y-%m-%d')
        return str(value)

    def process_excel_data(self, data):
        """ Converts pywintypes.datetime to strings for all values in the extracted data. """
        processed_data = []
        for row in data:
            processed_data.append([self.convert_pywintypes_to_string(value) for value in row])
        return processed_data

    def process_scheduling_files(self):
        """ Extract data from 'Indy Scheduling tool' and 'SB Scheduling Tool', combine all workbook data, and save once. """
        all_combined_data = []  # To store all data from all workbooks

        try:
            logging.debug(f"Base path: {self.base_path}")
            if not os.path.exists(self.base_path):
                logging.error(f"Base path does not exist: {self.base_path}")
                print(f"Base path does not exist: {self.base_path}")  # Debugging print
                return None

            try:
                # Use DispatchEx to create a new Excel instance
                excel = win32.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                logging.debug("Excel application started successfully.")
            except Exception as e:
                logging.error(f"Failed to create Excel application: {e}")
                print(f"Failed to create Excel application: {e}")  # Debugging print
                return None

            workbooks = {}
            try:
                # Open all workbooks in the base path
                for filename, required_subdir in self.files_info.items():
                    file_path = self.find_file(self.base_path, filename)
                    if file_path is None:
                        logging.warning(f"File not found: {filename}")
                        print(f"File not found: {filename}")  # Debugging print
                        continue
                    else:
                        # print(f"File found: {file_path}")  # Debugging print
                        logging.info(f"File found: {file_path}")

                    # Open the workbook and list sheets
                    try:
                        wb = excel.Workbooks.Open(file_path, False, True, None, self.password, '', True)
                        sheets = [sheet.Name for sheet in wb.Sheets]  # List all sheet names
                        # print(f"Sheets found in {filename}: {sheets}")  # Debugging print

                        # Process 'Indy Scheduling tool' and 'SB Scheduling Tool'
                        if 'Indy Scheduling tool' in sheets and 'SB Scheduling Tool' in sheets:
                            indy_ws = wb.Sheets("Indy Scheduling tool")
                            indy_data = self.process_excel_data(indy_ws.UsedRange.Value)  # Convert data
                            indy_scheduling_df = pd.DataFrame(indy_data[1:], columns=indy_data[0])

                            sb_ws = wb.Sheets("SB Scheduling Tool")
                            sb_data = self.process_excel_data(sb_ws.UsedRange.Value)  # Convert data
                            sb_scheduling_df = pd.DataFrame(sb_data[1:], columns=sb_data[0])
                            
                            # Convert column names to strings
                            indy_scheduling_df.columns = indy_scheduling_df.columns.astype(str)
                            sb_scheduling_df.columns = sb_scheduling_df.columns.astype(str)

                            # Set unnamed columns to row 3 values if applicable
                            indy_scheduling_df.columns = [
                                col if not col.startswith('Unnamed') else indy_scheduling_df.iloc[1, i] for i, col in enumerate(indy_scheduling_df.columns)
                            ]
                            sb_scheduling_df.columns = [
                                col if not col.startswith('Unnamed') else sb_scheduling_df.iloc[1, i] for i, col in enumerate(sb_scheduling_df.columns)
                            ]

                            # Rename only the first 3 columns to 'ABS', 'Patient', 'Caregiver'
                            indy_scheduling_df.columns = ['ABS','Patient', 'Caregiver'] + list(indy_scheduling_df.columns[3:])
                            sb_scheduling_df.columns = ['ABS','Patient', 'Caregiver'] + list(sb_scheduling_df.columns[3:])

                            # Drop columns D through L (index 3 to 11)
                            indy_scheduling_df_cleaned = indy_scheduling_df.drop(indy_scheduling_df.columns[3:12], axis=1)
                            sb_scheduling_df_cleaned = sb_scheduling_df.drop(sb_scheduling_df.columns[3:12], axis=1)

                            # Replace non-numeric values and blanks with 0 in columns D onwards
                            indy_scheduling_df_cleaned.iloc[:, 3:] = indy_scheduling_df_cleaned.iloc[:, 3:].apply(pd.to_numeric, errors='coerce').fillna(0)
                            sb_scheduling_df_cleaned.iloc[:, 3:] = sb_scheduling_df_cleaned.iloc[:, 3:].apply(pd.to_numeric, errors='coerce').fillna(0)

                            # Append the cleaned DataFrames to the all_combined_data list
                            all_combined_data.append(indy_scheduling_df_cleaned)
                            all_combined_data.append(sb_scheduling_df_cleaned)

                        else:
                            print(f"Required sheets not found in {filename}")

                    except Exception as e:
                        logging.error(f"Error opening file {file_path}: {e}")
                        print(f"Error opening file {file_path}: {e}")  # Debugging print
                        continue

            finally:
                # Close all workbooks and Excel application
                for wb in workbooks.values():
                    wb.Close(False)
                excel.Quit()
                del excel
                logging.debug("Excel application closed.")

            # Combine all data into a single DataFrame
            if all_combined_data:
                df = pd.concat(all_combined_data, ignore_index=True)
                
                # Remove unnecessary rows like Active Patient Count or Service Count
                df_cleaned = df[~df['ABS'].str.contains("Active Patient Count|Service Count|Med Rec", na=False)].copy()
                
                # Replace 'None' and empty strings with NaN to ensure proper forward filling
                df_cleaned.replace(to_replace=[None, 'None', ''], value=np.nan, inplace=True)

                # Forward fill the merged cells in 'ABS' and 'Patient' using .loc[]
                df_cleaned.iloc[:, :2] = df_cleaned.iloc[:, :2].ffill()
                
                # Drop rows where the 'Caregiver' column is NaN (i.e., no value in the Caregiver column)
                df_cleaned = df_cleaned.dropna(subset=['Caregiver'])

                # Export the combined DataFrame to a CSV file
                # df_cleaned.to_csv(output_file, index=False)
                # print(f"Data has been saved to: {output_file}")
                return df_cleaned
            else:
                # print("No data was extracted from the workbooks.")
                logging.warning("No data was extracted from any of the workbooks.")
                return None

        except Exception as e:
            logging.error(f"Error in process_scheduling_files: {e}")
            # print(f"Error: {e}")

# Run the extraction process
if __name__ == "__main__":
    extractor = EmployeeRecordsExtractor()
    df = extractor.process_scheduling_files()
    if df is not None:
        # Output DataFrame in CSV format without index
        print(df.to_csv(index=False))
    else:
        print("No eligible patient data was extracted.")
