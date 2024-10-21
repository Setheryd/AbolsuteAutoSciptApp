import sys
import pandas as pd
from billing_files_extractor import BillingFilesDataExtractor

def main():
    extractor = BillingFilesDataExtractor()
    
    # Extract the DataFrame without using output_to_csv
    try:
        df = extractor.process_billing_files()  # Removed the output_to_csv argument
    except pd.errors.ParserError as e:
        print(f"Error parsing data: {e}")
        sys.exit(1)
    except pd.errors.EmptyDataError as e:
        print(f"No data found in the file: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

    if df is not None:
        # Ensure DataFrame is not empty
        # Define the values to search for
        keywords = ['ATTC', 'HMK', 'PERS', 'NUTS', 'CHOICE', 'IHCC', 'SFC']
        excluded_terms = ['Switched Payer', 'Discharged']

        # Filter out rows containing 'Switched Payer' or 'Discharged' in the Medical Record Number column
        filtered_data = df[~df['Medical Record Number'].str.contains('|'.join(excluded_terms), na=False)]

        # Count the occurrences of each keyword in the Medical Record Number column
        occurrences = {keyword: filtered_data['Medical Record Number'].str.contains(keyword).sum() for keyword in keywords}

        occurrences_df = pd.DataFrame(list(occurrences.items()), columns=['Service', 'Count'])
        
        csv_output = occurrences_df.to_csv(index=False)

        print(csv_output)

    else:
        sys.exit(1)  # Exit with error code if no data

if __name__ == "__main__":
    main()
