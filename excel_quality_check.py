import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Tuple

class ExcelQualityChecker:
    def __init__(self, file_path: str):
        """Initialize the Excel Quality Checker.
        
        Args:
            file_path (str): Path to the Excel file
        """
        self.file_path = Path(file_path)
        self.df = pd.read_excel(file_path)
        self.report = {}

    def check_null_values(self) -> Dict:
        """Check for null values in each column."""
        null_counts = self.df.isnull().sum()
        null_percentages = (null_counts / len(self.df)) * 100
        
        self.report['null_values'] = {
            'counts': null_counts.to_dict(),
            'percentages': null_percentages.to_dict()
        }
        return self.report['null_values']

    def check_duplicates(self) -> Dict:
        """Check for duplicate rows in the dataset."""
        duplicates = self.df.duplicated()
        duplicate_count = duplicates.sum()
        duplicate_percentage = (duplicate_count / len(self.df)) * 100
        
        self.report['duplicates'] = {
            'total_count': int(duplicate_count),
            'percentage': float(duplicate_percentage),
            'duplicate_rows': self.df[duplicates].index.tolist()
        }
        return self.report['duplicates']

    def check_column_types(self) -> Dict:
        """Check data types of each column."""
        self.report['column_types'] = self.df.dtypes.astype(str).to_dict()
        return self.report['column_types']

    def run_all_checks(self) -> Dict:
        """Run all available checks and return the complete report."""
        self.check_null_values()
        self.check_duplicates()
        self.check_column_types()
        return self.report

def main():
    """Example usage of the ExcelQualityChecker."""
    try:
        # Replace with your Excel file path
        checker = ExcelQualityChecker("your_file.xlsx")
        report = checker.run_all_checks()
        
        print("\n=== Excel Data Quality Report ===")
        
        print("\nNull Values Summary:")
        for col, stats in report['null_values']['counts'].items():
            print(f"{col}: {stats} nulls ({report['null_values']['percentages'][col]:.2f}%)")
        
        print("\nDuplicates Summary:")
        print(f"Total duplicate rows: {report['duplicates']['total_count']}")
        print(f"Percentage of duplicates: {report['duplicates']['percentage']:.2f}%")
        
        print("\nColumn Types:")
        for col, dtype in report['column_types'].items():
            print(f"{col}: {dtype}")

    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()