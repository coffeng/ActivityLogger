"""
Debug script to check key format differences between summary and log files
"""
import csv
import os

def check_key_formats():
    """Check key formats in both files"""
    log_file = "ActivityLogger.csv"
    summary_file = "ActivitySummary.csv"
    
    print("=== Key Format Analysis ===")
    
    # Check ActivityLogger.csv keys
    if os.path.exists(log_file):
        print(f"\nKeys in {log_file}:")
        with open(log_file, 'r', encoding='utf-8') as f:
            reader = list(csv.reader(f))
            if len(reader) > 1:
                headers = reader[0]
                if 'ApplicationKey' in headers:
                    app_key_index = headers.index('ApplicationKey')
                    unique_keys = set()
                    for row in reader[1:]:
                        if len(row) > app_key_index:
                            unique_keys.add(row[app_key_index])
                    
                    for key in sorted(unique_keys):
                        print(f"  - '{key}'")
    
    # Check ActivitySummary.csv keys
    if os.path.exists(summary_file):
        print(f"\nKeys in {summary_file}:")
        with open(summary_file, 'r', encoding='utf-8') as f:
            reader = list(csv.reader(f))
            for row in reader:
                if row and not row[0].startswith('#') and row[0]:
                    if row[0] != 'Key':  # Skip header
                        print(f"  - '{row[0]}'")

if __name__ == "__main__":
    check_key_formats()