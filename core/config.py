"""
Configuration management for Activity Logger
"""
import os
import csv
import datetime
from collections import defaultdict
from .utils import format_duration


class ConfigManager:
    """Manages application configuration and category settings"""
    
    def __init__(self, log_path):
        self.log_path = log_path
        self.config_path = os.path.join(os.path.dirname(log_path), "ActivitySummary.csv")
        self.default_categories = {
            "excel": "Work - Office",
            "winword": "Work - Office", 
            "powerpnt": "Work - Office",
            "outlook": "Email",
            "chrome": "Web Browsing",
            "firefox": "Web Browsing",
            "edge": "Web Browsing",
            "teams": "Meetings",
            "slack": "Communication",
            "zoom": "Meetings",
            "notepad": "Notes",
            "code": "Development",
            "pycharm": "Development",
            "cmd": "Terminal",
            "powershell": "Terminal"
        }

    def load_app_categories(self):
        """Load app categories from ActivitySummary.csv file"""
        if not os.path.exists(self.config_path):
            # Create initial CSV file
            self.save_app_categories(self.default_categories, {}, {})
            return self.default_categories
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                
                # Skip comment lines and find headers
                headers = None
                for row in reader:
                    if row and not row[0].startswith('#') and row[0]:
                        headers = row
                        break
                
                if not headers or 'Key' not in headers or 'Category' not in headers:
                    print("Invalid summary file format, using defaults")
                    self.save_app_categories(self.default_categories, {}, {})
                    return self.default_categories
                
                # Find column indices
                key_idx = headers.index('Key')
                category_idx = headers.index('Category')
                
                # Extract key-category mapping
                categories = {}
                for row in reader:
                    if len(row) > max(key_idx, category_idx):
                        key = row[key_idx].strip()
                        category = row[category_idx].strip()
                        if key and category:
                            categories[key] = category
            
            if categories:
                print(f"Loaded {len(categories)} categories from ActivitySummary.csv")
                # Merge with defaults for any missing keys
                merged_categories = self.default_categories.copy()
                merged_categories.update(categories)
                return merged_categories
            else:
                print("No valid categories found in summary file, using defaults")
                self.save_app_categories(self.default_categories, {}, {})
                return self.default_categories
                
        except Exception as e:
            print(f"Error loading ActivitySummary.csv: {e}")
            return self.default_categories

    def save_app_categories(self, categories=None, row_counts=None, durations=None):
        """Save app categories to ActivitySummary.csv file with statistics"""
        if categories is None:
            categories = self.default_categories
            
        # Get current statistics from log file
        if row_counts is None or durations is None:
            row_counts, durations = self.calculate_category_stats(categories)
        
        # Find all keys that appear in the log file
        discovered_keys = self._discover_keys_from_log()
        
        # Merge existing categories with discovered keys
        expanded_categories = categories.copy()
        
        # Add discovered keys that aren't already categorized
        for key in discovered_keys:
            if key not in expanded_categories:
                expanded_categories[key] = self._auto_categorize_key(key)
        
        # Create CSV data, but only include keys with non-zero counts
        csv_rows = []
        for key, category in expanded_categories.items():
            count = row_counts.get(key, 0)
            duration_seconds = durations.get(key, 0)
            
            # Only include keys that have been used (non-zero count)
            if count > 0:
                csv_rows.append({
                    'key': key,
                    'category': category,
                    'count': count,
                    'duration': format_duration(duration_seconds),
                    'duration_seconds': duration_seconds  # For sorting
                })
        
        # Sort by duration (highest first)
        csv_rows.sort(key=lambda x: x['duration_seconds'], reverse=True)
        
        # Write CSV file
        self._write_csv_file(csv_rows)

    def _discover_keys_from_log(self):
        """Find all keys that appear in the log file"""
        discovered_keys = set()
        if not os.path.exists(self.log_path):
            return discovered_keys
            
        try:
            with open(self.log_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader, None)
                if headers:
                    try:
                        process_idx = headers.index('ProcessName')
                        window_idx = headers.index('WindowTitle') 
                        details_idx = headers.index('WindowDetails')
                        
                        for row in reader:
                            if len(row) > max(process_idx, window_idx, details_idx):
                                process_name = row[process_idx].lower()
                                window_title = row[window_idx].lower()
                                window_details = row[details_idx].lower()
                                
                                # Extract potential keys from process names
                                if process_name.endswith('.exe'):
                                    base_name = process_name[:-4]
                                    discovered_keys.add(base_name)
                                
                                # Extract potential keys from window titles
                                words = window_title.split()
                                for word in words:
                                    if len(word) > 3:
                                        discovered_keys.add(word.lower())
                                
                                # Extract potential keys from window details
                                words = window_details.split()
                                for word in words:
                                    if len(word) > 3:
                                        discovered_keys.add(word.lower())
                                        
                    except (ValueError, IndexError):
                        pass
        except Exception as e:
            print(f"Error discovering keys: {e}")
            
        return discovered_keys

    def _auto_categorize_key(self, key):
        """Automatically categorize a key based on common patterns"""
        if any(x in key for x in ['excel', 'word', 'powerpoint', 'office']):
            return "Work - Office"
        elif any(x in key for x in ['chrome', 'firefox', 'edge', 'browser']):
            return "Web Browsing"
        elif any(x in key for x in ['teams', 'zoom', 'meet', 'skype']):
            return "Meetings"
        elif any(x in key for x in ['outlook', 'mail', 'thunderbird']):
            return "Email"
        elif any(x in key for x in ['code', 'studio', 'pycharm', 'eclipse']):
            return "Development"
        elif any(x in key for x in ['cmd', 'powershell', 'terminal', 'bash']):
            return "Terminal"
        elif any(x in key for x in ['notepad', 'note', 'text']):
            return "Notes"
        elif any(x in key for x in ['slack', 'discord', 'chat']):
            return "Communication"
        else:
            return "Uncategorized"

    def _write_csv_file(self, csv_rows):
        """Write the CSV file with metadata and data"""
        try:
            with open(self.config_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # Write header with timestamp
                writer.writerow(['# ActivitySummary.csv'])
                writer.writerow(['# Last Updated:', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                writer.writerow(['# Total Entries:', len(csv_rows)])
                writer.writerow([])  # Empty row for separation
                
                # Write CSV headers
                writer.writerow(['Key', 'Category', 'Count', 'Duration'])
                
                # Write data rows
                for row in csv_rows:
                    writer.writerow([
                        row['key'],
                        row['category'], 
                        row['count'],
                        row['duration']
                    ])
                    
        except Exception as e:
            print(f"Error saving ActivitySummary.csv: {e}")

    def calculate_category_stats(self, categories):
        """Calculate row counts and durations for each key from log file"""
        row_counts = defaultdict(int)
        durations = defaultdict(int)
        
        if not os.path.exists(self.log_path):
            return row_counts, durations
        
        try:
            with open(self.log_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader, None)
                if not headers:
                    return row_counts, durations
                
                # Find column indices
                try:
                    process_idx = headers.index('ProcessName')
                    window_idx = headers.index('WindowTitle') 
                    details_idx = headers.index('WindowDetails')
                    duration_idx = headers.index('DurationSeconds')
                except ValueError:
                    return row_counts, durations
                
                for row in reader:
                    if len(row) > max(process_idx, window_idx, details_idx, duration_idx):
                        process_name = row[process_idx].lower()
                        window_title = row[window_idx].lower()
                        window_details = row[details_idx].lower()
                        
                        try:
                            duration_seconds = int(row[duration_idx])
                        except (ValueError, IndexError):
                            continue
                        
                        # Check all possible keys
                        all_keys = set(categories.keys())
                        
                        # Add discovered keys from current row
                        if process_name.endswith('.exe'):
                            all_keys.add(process_name[:-4])
                        
                        # Check each key for matches
                        for key in all_keys:
                            if (key in process_name or 
                                key in window_title or 
                                key in window_details):
                                row_counts[key] += 1
                                durations[key] += duration_seconds
                                break  # Only count once per row
                                
        except Exception as e:
            print(f"Error calculating stats: {e}")
            
        return row_counts, durations