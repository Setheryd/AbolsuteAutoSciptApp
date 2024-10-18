# extract_patients.py
import sys
from patient_data_extractor import PatientDataExtractor

def main():
    extractor = PatientDataExtractor()
    df = extractor.extract_eligible_patients()  # Output to stdout by default
    if df is not None:
        # DataFrame is already printed to stdout by extract_eligible_patients
        print(df.to_csv(index=False))
        pass
    else:
        sys.exit(1)  # Exit with error code if no data

if __name__ == "__main__":
    main()
