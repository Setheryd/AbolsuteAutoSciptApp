import win32com.client as win32
import pandas as pd
import os
from datetime import datetime
import sys

class PatientDataExtractor:
    def __init__(self, base_path=None, password="abs$1018$B"):
        self.base_path = base_path or f"C:\\Users\\{os.getlogin()}\\OneDrive - Ability Home Health, LLC\\"
        self.password = password
        self.files_info = {
            "Absolute Patient Records.xlsm": "Absolute Operation",
            "Absolute Patient Records IHCC.xlsm": "IHCC",
            "Absolute Patient Records PERS.xlsm": "IHCC"
        }

    def find_file(self, base_path, filename, max_depth=5):
        """
        Optimized search for a file in a directory up to a specified depth using os.scandir for speed.
        """
        def scan_directory(path, current_depth):
            if current_depth > max_depth:
                return None
            try:
                with os.scandir(path) as it:
                    for entry in it:
                        if entry.is_file() and entry.name.lower() == filename.lower():
                            return entry.path
                        elif entry.is_dir():
                            found_file = scan_directory(entry.path, current_depth + 1)
                            if found_file:
                                return found_file
            except PermissionError:
                return None
            except Exception:
                return None

        return scan_directory(base_path, 0)

    def extract_eligible_patients(self):
        """
        Extract eligible patients from Excel files.

        Returns:
            pd.DataFrame or None: The extracted DataFrame, or None if no data was collected.
        """
        if not os.path.exists(self.base_path):
            print(f"Base path does not exist: {self.base_path}")
            return None

        collected_data = []
        excel = None
        workbooks = {}
        try:
            # Initialize Excel application
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.ScreenUpdating = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False
            excel.AlertBeforeOverwriting = False

            # Open all workbooks
            for filename, required_subdir in self.files_info.items():
                file_path = self.find_file(self.base_path, filename)
                if file_path is None:
                    continue

                # Verify the file is in the correct subdirectory
                normalized_path = os.path.normpath(file_path)
                if required_subdir.lower() not in normalized_path.lower():
                    continue

                try:
                    wb = excel.Workbooks.Open(
                        file_path,            # Filename
                        False,                # UpdateLinks
                        True,                 # ReadOnly
                        None,                 # Format
                        self.password,        # Password
                        '',                   # WriteResPassword
                        True,                 # IgnoreReadOnlyRecommended
                        None,                 # Origin
                        None,                 # Delimiter
                        False,                # Editable
                        False,                # Notify
                        None,                 # Converter
                        False,                # AddToMru
                        False,                # Local
                        0                     # CorruptLoad
                    )
                    workbooks[filename] = wb
                except Exception:
                    continue

            # Process each workbook
            for filename, wb in workbooks.items():
                try:
                    ws = wb.Sheets("Patient Information")
                    used_range = ws.UsedRange
                    data = used_range.Value  # Read all data at once

                    if not data:
                        continue

                    # Assuming data is a tuple of tuples
                    data_rows = data[1:]  # Skip header row

                    for row in data_rows:
                        if filename == "Absolute Patient Records.xlsm":
                            # For "Absolute Patient Records.xlsm", use column J (10th column) for NOA start date
                            noa_start_date = row[9] if len(row) >= 10 else None  # Column J (10th)
                            discharge_date = row[11] if len(row) >= 12 else None  # Column L (12th)
                            if noa_start_date:
                                patient_name = row[2] if len(row) >= 3 else ''  # Column C (3rd)
                                collected_data.append({
                                    "Patient Name": patient_name,
                                    "First NOA Date": noa_start_date,
                                    "Discharge Date": discharge_date
                                })
                        else:
                            # For other files, use column K (11th column) for NOA start date
                            noa_start_date = row[10] if len(row) >= 11 else None  # Column K (11th)
                            discharge_date = row[13] if len(row) >= 14 else None  # Column N (14th)
                            if noa_start_date:
                                patient_name = row[3] if len(row) >= 4 else ''  # Column D (4th)
                                collected_data.append({
                                    "Patient Name": patient_name,
                                    "First NOA Date": noa_start_date,
                                    "Discharge Date": discharge_date
                                })
                except Exception:
                    continue

        except Exception:
            return None
        finally:
            # Close all workbooks and quit Excel
            if workbooks:
                for wb in workbooks.values():
                    try:
                        wb.Close(False)
                    except Exception:
                        pass
            if excel:
                excel.Quit()
                del excel

        if collected_data:
            # Create a DataFrame from the collected data
            df = pd.DataFrame(collected_data)

            # Convert date columns to strings in 'YYYY-MM-DD' format
            def format_date(date_obj):
                if isinstance(date_obj, datetime):
                    return date_obj.strftime('%Y-%m-%d')
                elif isinstance(date_obj, str):
                    try:
                        # Attempt to parse the string to a datetime object
                        parsed_date = pd.to_datetime(date_obj, errors='coerce')
                        if pd.notnull(parsed_date):
                            return parsed_date.strftime('%Y-%m-%d')
                        else:
                            return date_obj  # Return as is if parsing fails
                    except Exception:
                        return date_obj  # Return as is if any exception occurs
                else:
                    return date_obj  # Return as is for other types

            df['First NOA Date'] = df['First NOA Date'].apply(format_date)
            df['Discharge Date'] = df['Discharge Date'].apply(format_date)

            # Drop duplicates based on all columns
            df.drop_duplicates(inplace=True)

            return df
        else:
            print("No data collected.")
            return None

def main():
    extractor = PatientDataExtractor()
    df = extractor.extract_eligible_patients()
    if df is not None:
        # Output DataFrame in CSV format without index
        print(df.to_csv(index=False))
    else:
        print("No eligible patient data was extracted.")

if __name__ == "__main__":
    main()