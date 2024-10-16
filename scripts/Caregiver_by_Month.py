import sys
import pandas as pd
from caregiver_data_extractor import CaregiverDataExtractor

def main():
    extractor = CaregiverDataExtractor()
    
    # Extract the DataFrame without exporting to CSV
    df = extractor.extract_caregivers(output_to_csv=False)

    if df is not None:
        # Print the shape and the first few rows of the DataFrame for debugging
        print("DataFrame Shape:", df.shape)
        print("DataFrame Head:\n", df.head())

        if df.empty:
            print("The DataFrame is empty.")
            sys.exit(1)  # Exit with error code if no data

        # Convert 'Date of Hire' and 'Term Date' to datetime for proper handling
        df["Date of Hire (H)"] = pd.to_datetime(df["Date of Hire (H)"], errors='coerce')
        df["Term Date (J)"] = pd.to_datetime(df["Term Date (J)"], errors='coerce')

        # Determine the date range (from the earliest month to the current month)
        start_date = df["Date of Hire (H)"].min().replace(day=1)
        end_date = pd.Timestamp.today().replace(day=1)

        # Create a list of months between the start and end date
        all_months = pd.date_range(start=start_date, end=end_date, freq='MS')

        # Create an empty list to store the count of active caregivers per month
        active_patient_counts = []

        # Loop through each month and count the number of active caregivers
        for month in all_months:
            active_patients = df[
                (df["Date of Hire (H)"] <= month) &
                ((df["Term Date (J)"].isna()) | (df["Term Date (J)"] >= month))
            ]
            active_patient_counts.append({'Month': month.strftime('%B %Y'), 'Active Caregivers': len(active_patients)})

        # Convert the list of results into a DataFrame
        df_active_patients_by_month = pd.DataFrame(active_patient_counts)

        # Print the result DataFrame for debugging
        print(df_active_patients_by_month.to_csv(index=False))

    else:
        print("No data was returned by the extractor.")
        sys.exit(1)  # Exit with error code if no data

if __name__ == "__main__":
    main()
