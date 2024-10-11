import sys
import pandas as pd
from caregiver_data_extractor import CaregiverDataExtractor

def main():
    extractor = CaregiverDataExtractor()
    
    # Extract the DataFrame without exporting to CSV
    df = extractor.extract_caregivers(output_to_csv=False)  # Ensure it doesn't output directly to stdout

    if df is not None:
        # Convert 'First NOA Date' and 'Discharge Date' to datetime for proper handling
        df["Date of Hire (H)"] = pd.to_datetime(df["Date of Hire (H)"], errors='coerce')
        df["Term Date (J)"] = pd.to_datetime(df["Term Date (J)"], errors='coerce')

        # Determine the date range (from the earliest month to the current month)
        start_date = df["Date of Hire (H)"].min().replace(day=1)
        end_date = pd.Timestamp.today().replace(day=1)

        # Create a list of months between the start and end date
        all_months = pd.date_range(start=start_date, end=end_date, freq='MS')

        # Create an empty list to store the count of active patients per month
        active_patient_counts = []

        # Loop through each month and count the number of active patients
        for month in all_months:
            active_patients = df[
                (df["Date of Hire (H)"] <= month) &
                ((df["Term Date (J)"].isna()) | (df["Term Date (J)"] >= month))
            ]
            active_patient_counts.append({'Month': month.strftime('%B %Y'), 'Active Patients': len(active_patients)})

        # Convert the list of results into a DataFrame
        df_active_patients_by_month = pd.DataFrame(active_patient_counts)

        # Output the DataFrame to stdout
        print(df_active_patients_by_month.to_csv(index=False))

    else:
        sys.exit(1)  # Exit with error code if no data

if __name__ == "__main__":
    main()
