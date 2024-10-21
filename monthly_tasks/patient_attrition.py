# churn_attrition_analysis.py

import pandas as pd
import os
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import MaxNLocator
from data_extraction.patient_data_extractor import PatientDataExtractor


class ChurnAttritionAnalyzer:
    def __init__(self, extractor: PatientDataExtractor):
        self.extractor = extractor
        self.report_filename = "monthly_report.csv"
        self.outliers_filename = "outliers_report.xlsx"

    def load_data(self):
        """
        Load the patient data using the extractor.

        Returns:
            pd.DataFrame or None
        """
        df = self.extractor.extract_eligible_patients()
        if df is not None:
            # Convert date columns from strings to datetime objects
            df["First NOA Date"] = pd.to_datetime(df["First NOA Date"], errors="coerce")
            df["Discharge Date"] = pd.to_datetime(df["Discharge Date"], errors="coerce")
            # Drop records with invalid First NOA Date
            df = df.dropna(subset=["First NOA Date"])
            return df
        else:
            print("No eligible patient data was extracted.")
            return None

    def generate_monthly_report(self, df, report_month=None):
        """
        Generate a churn and attrition rate report for a specific month.

        Args:
            df (pd.DataFrame): DataFrame containing patient data.
            report_month (datetime, optional): The month for which to generate the report. Defaults to last month.

        Returns:
            dict: A dictionary containing the report metrics.
        """
        if report_month is None:
            today = datetime.today()
            report_month = pd.to_datetime(
                datetime(today.year, today.month, 1)
            ) - pd.DateOffset(months=1)
        else:
            report_month = pd.to_datetime(report_month)

        # Define the start and end of the report month
        start_of_month = report_month.replace(day=1)
        end_of_month = (start_of_month + pd.DateOffset(months=1)) - pd.DateOffset(
            days=1
        )

        # Ensure start_of_month and end_of_month are offset-naive
        if start_of_month.tzinfo is not None:
            start_of_month = start_of_month.tz_localize(None)
        if end_of_month.tzinfo is not None:
            end_of_month = end_of_month.tz_localize(None)

        # Filter active patients at the start and end of the month
        active_start = df[
            (df["First NOA Date"] <= start_of_month)
            & ((df["Discharge Date"].isna()) | (df["Discharge Date"] > start_of_month))
        ]
        active_end = df[
            (df["First NOA Date"] <= end_of_month)
            & ((df["Discharge Date"].isna()) | (df["Discharge Date"] > end_of_month))
        ]

        # Calculate the number of active patients
        active_start_count = active_start["Patient Name"].nunique()
        active_end_count = active_end["Patient Name"].nunique()

        # Calculate net change
        net_change = active_end_count - active_start_count

        # Calculate active patients per month for average net change
        months = pd.period_range(
            df["First NOA Date"].min().to_period("M"),
            end_of_month.to_period("M"),
            freq="M",
        )
        active_counts = []
        for month in months:
            month_start = month.to_timestamp()
            month_end = (month_start + pd.DateOffset(months=1)) - pd.DateOffset(days=1)
            active = df[
                (df["First NOA Date"] <= month_end)
                & ((df["Discharge Date"].isna()) | (df["Discharge Date"] > month_end))
            ]
            active_counts.append(active["Patient Name"].nunique())

        # Calculate monthly net changes
        net_changes = [
            active_counts[i] - active_counts[i - 1]
            for i in range(1, len(active_counts))
        ]
        average_net_change = pd.Series(net_changes).mean() if net_changes else 0

        # Calculate churn and attrition rates
        churn = (
            (abs(net_change) / active_start_count) * 100 if active_start_count else 0
        )
        # Attrition rate: number of discharges during the month divided by average number of patients
        discharges = df[
            (df["Discharge Date"] >= start_of_month)
            & (df["Discharge Date"] <= end_of_month)
        ]
        num_discharges = discharges["Patient Name"].nunique()
        average_patients = (active_start_count + active_end_count) / 2
        attrition = (num_discharges / average_patients) * 100 if average_patients else 0

        # Prepare report
        report = {
            "Report Month": report_month.strftime("%B %Y"),
            "Starting Patient Count": active_start_count,
            "Ending Patient Count": active_end_count,
            "Net Change": net_change,
            "Average Net Change": round(average_net_change, 2),
            "Churn Rate (%)": round(churn, 2),
            "Attrition Rate (%)": round(attrition, 2),
        }

        return report

    def generate_all_monthly_reports(self, df):
        """
        Generate monthly reports for all months in the data.

        Args:
            df (pd.DataFrame): DataFrame containing patient data.

        Returns:
            pd.DataFrame: DataFrame containing all monthly reports.
        """
        months = pd.period_range(
            df["First NOA Date"].min().to_period("M"),
            pd.Timestamp.today().to_period("M"),
            freq="M",
        )
        report_data = []
        for month in months:
            report = self.generate_monthly_report(df, report_month=month.to_timestamp())
            report_data.append(report)
        report_df = pd.DataFrame(report_data)
        return report_df

    def detect_outliers(self, df_reports, column, method="IQR"):
        """
        Detect outliers in a specific column of the reports DataFrame.

        Args:
            df_reports (pd.DataFrame): DataFrame containing monthly reports.
            column (str): The column in which to detect outliers.
            method (str, optional): The method to use for outlier detection ('IQR' or 'Z-score'). Defaults to 'IQR'.

        Returns:
            pd.DataFrame: DataFrame containing outlier records.
        """
        if method == "IQR":
            Q1 = df_reports[column].quantile(0.25)
            Q3 = df_reports[column].quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - 1.5 * IQR
            upper_bound = Q3 + 1.5 * IQR
            outliers = df_reports[
                (df_reports[column] < lower_bound) | (df_reports[column] > upper_bound)
            ]
        elif method == "Z-score":
            z_scores = stats.zscore(df_reports[column].dropna())
            outliers = df_reports.iloc[(z_scores < -3) | (z_scores > 3)]
        else:
            raise ValueError("Unsupported method. Use 'IQR' or 'Z-score'.")

        return outliers

    def save_report(self, report_df, filename=None):
        """
        Save the report DataFrame to a CSV file.

        Args:
            report_df (pd.DataFrame): DataFrame containing monthly reports.
            filename (str, optional): The filename for the CSV. Defaults to 'monthly_report.csv'.
        """
        if filename is None:
            filename = self.report_filename
        report_df.to_csv(filename, index=False)
        print(f"Monthly reports have been saved to {filename}.")

    def save_outliers(self, outliers_churn, outliers_attrition, filename=None):
        """
        Save the outliers to an Excel file.

        Args:
            outliers_churn (pd.DataFrame): DataFrame containing churn rate outliers.
            outliers_attrition (pd.DataFrame): DataFrame containing attrition rate outliers.
            filename (str, optional): The filename for the Excel file. Defaults to 'outliers_report.xlsx'.
        """
        if filename is None:
            filename = self.outliers_filename
        with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
            if not outliers_churn.empty:
                outliers_churn[["Report Month", "Churn Rate (%)"]].to_excel(
                    writer, sheet_name="Churn Rate Outliers", index=False
                )
            if not outliers_attrition.empty:
                outliers_attrition[["Report Month", "Attrition Rate (%)"]].to_excel(
                    writer, sheet_name="Attrition Rate Outliers", index=False
                )
        print(f"Outliers have been saved to {filename}.")

    def generate_charts(self, report_df, outliers_churn, outliers_attrition):
        """
        Generate charts showing churn and attrition rates over time.

        Args:
            report_df (pd.DataFrame): DataFrame containing monthly reports.
            outliers_churn (pd.DataFrame): DataFrame containing churn rate outliers.
            outliers_attrition (pd.DataFrame): DataFrame containing attrition rate outliers.
        """
        # Convert 'Report Month' to datetime for plotting
        report_df["Report Month Date"] = pd.to_datetime(
            report_df["Report Month"], format="%B %Y"
        )

        # Plot Churn and Attrition Rates
        plt.figure(figsize=(14, 7))
        plt.plot(
            report_df["Report Month Date"],
            report_df["Churn Rate (%)"],
            marker="o",
            label="Churn Rate (%)",
        )
        plt.plot(
            report_df["Report Month Date"],
            report_df["Attrition Rate (%)"],
            marker="o",
            label="Attrition Rate (%)",
        )

        # Highlight outliers
        if not outliers_churn.empty:
            outliers_churn_dates = pd.to_datetime(
                outliers_churn["Report Month"], format="%B %Y"
            )
            plt.scatter(
                outliers_churn_dates,
                outliers_churn["Churn Rate (%)"],
                color="red",
                label="Churn Rate Outliers",
                zorder=5,
            )
        if not outliers_attrition.empty:
            outliers_attrition_dates = pd.to_datetime(
                outliers_attrition["Report Month"], format="%B %Y"
            )
            plt.scatter(
                outliers_attrition_dates,
                outliers_attrition["Attrition Rate (%)"],
                color="orange",
                label="Attrition Rate Outliers",
                zorder=5,
            )

        plt.xlabel("Month")
        plt.ylabel("Rate (%)")
        plt.title("Churn and Attrition Rates Over Time")
        plt.legend()

        # Set major ticks to every 3 months
        ax = plt.gca()
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))  # Every 2 months
        ax.xaxis.set_major_formatter(
            mdates.DateFormatter("%B %Y")
        )  # e.g., January 2023

        # Rotate date labels for better readability
        plt.xticks(rotation=45)

        plt.tight_layout()
        plt.savefig("churn_attrition_rates.png")
        plt.show()

    def run_analysis(self):
        """
        Run the complete churn and attrition analysis.
        """
        df = self.load_data()
        if df is None:
            print("No data available for analysis.")
            return

        # Generate all monthly reports
        report_df = self.generate_all_monthly_reports(df)
        print("Monthly Report:")
        print(report_df.tail())  # Display the last few reports

        # Save the report to CSV
        self.save_report(report_df)

        # Detect outliers in Churn Rate and Attrition Rate
        outliers_churn = self.detect_outliers(report_df, "Churn Rate (%)", method="IQR")
        outliers_attrition = self.detect_outliers(
            report_df, "Attrition Rate (%)", method="IQR"
        )

        # Report outliers
        if not outliers_churn.empty:
            print("\n=== Churn Rate Outliers Detected ===")
            print(outliers_churn[["Report Month", "Churn Rate (%)"]])
        else:
            print("\nNo outliers detected in Churn Rate.")

        if not outliers_attrition.empty:
            print("\n=== Attrition Rate Outliers Detected ===")
            print(outliers_attrition[["Report Month", "Attrition Rate (%)"]])
        else:
            print("\nNo outliers detected in Attrition Rate.")

        # Save outliers to an Excel file
        self.save_outliers(outliers_churn, outliers_attrition)

        # Generate charts
        self.generate_charts(report_df, outliers_churn, outliers_attrition)


def main():
    extractor = PatientDataExtractor()
    analyzer = ChurnAttritionAnalyzer(extractor)
    analyzer.run_analysis()


if __name__ == "__main__":
    main()
