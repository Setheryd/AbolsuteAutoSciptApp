import sys
import pandas as pd
from patient_data_extractor import PatientDataExtractor

def main():
    extractor = PatientDataExtractor()
    
    # Extract the DataFrame directly
    df = extractor.extract_eligible_patients()  # Removed output_to_csv argument

    if df is not None:
        # Convert 'First NOA Date' and 'Discharge Date' to datetime for proper handling
        df['First NOA Date'] = pd.to_datetime(df['First NOA Date'], errors='coerce')
        df['Discharge Date'] = pd.to_datetime(df['Discharge Date'], errors='coerce')

        # Calculate the tenure duration as the difference between 'Discharge Date' and 'First NOA Date'
        df["Tenure Duration"] = df["Discharge Date"] - df["First NOA Date"]

        # Define bins for grouping every 50 days
        bins = list(range(0, df["Tenure Duration"].max().days + 50, 50))

        # Create a new column that categorizes the tenure duration into groups of 50 days
        df["Tenure Group"] = pd.cut(df["Tenure Duration"].dt.days, bins=bins)

        # Count the occurrences in each group
        tenure_group_counts = df["Tenure Group"].value_counts().sort_index()

        # Convert to a DataFrame for better presentation
        tenure_group_counts_df = pd.DataFrame(tenure_group_counts).reset_index()
        tenure_group_counts_df.columns = ['Tenure Group', 'Count']

        # Modify the labels to match the format 'x-y days'
        tenure_group_counts_df['Tenure Group'] = tenure_group_counts_df['Tenure Group'].apply(
            lambda x: f"{int(x.left)}-{int(x.right) - 1} days"
        )
        
        csv_output = tenure_group_counts_df.to_csv(index=False)
        
        print(csv_output)

    else:
        sys.exit(1)  # Exit with error code if no data

if __name__ == "__main__":
    main()
