"""
Log viewer window
"""
import os
import csv
import datetime
import tkinter as tk
from tkinter import ttk
from core.utils import format_duration, ExeVersionInfo
from .category_selector import CategorySelector
import matplotlib
matplotlib.use('Agg')  # Use a non-interactive backend for safety
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import sys
import pefile


class LogViewer:
    """Activity log viewer window"""
    
    _instances = {}  # Class variable to track open instances
    
    def __init__(self, log_path):
        self._last_log_mtime = None
        self._last_log_data = None
        # Check if an instance for this log_path already exists
        if log_path in LogViewer._instances:
            existing_viewer = LogViewer._instances[log_path]
            try:
                if existing_viewer.root.winfo_exists():
                    existing_viewer.root.lift()
                    existing_viewer.root.focus_force()
                    return
            except:
                # Remove stale reference
                del LogViewer._instances[log_path]

        self.log_path = log_path
        self.summary_path = os.path.join(os.path.dirname(log_path), "ActivitySummary.csv")
        self.is_duplicate = log_path in LogViewer._instances
        LogViewer._instances[log_path] = self

        # Sorting state
        self.sort_column = None
        self.sort_reverse = False
        self.summary_sort_column = None
        self.summary_sort_reverse = False

        # Create main window
        self.root = tk.Tk()
        version_info = ExeVersionInfo()
        version = version_info.get_version()
        build_date = version_info.get_build_date()
        build_time = version_info.get_build_time()
        self.root.title(
            f"Activity Log Viewer - {os.path.basename(log_path)} | Build {version} {build_date} {build_time}"
        )
        self.root.geometry("1200x700")
        
        # Handle window close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Make tab buttons 150% in height
        style = ttk.Style(self.root)
        style.configure("TNotebook.Tab", minheight=36)

        # Create tabs
        self.activity_frame = ttk.Frame(self.notebook)
        self.summary_frame = ttk.Frame(self.notebook)
        self.graph_frame = ttk.Frame(self.notebook)  # New tab for stacked bar graph
        
        # When adding tabs to the notebook, add 2 spaces before and after the tab label
        self.notebook.add(self.activity_frame, text="  Activity Log  ")
        self.notebook.add(self.summary_frame, text="  Summary  ")
        self.notebook.add(self.graph_frame, text="  Category Graph  ")  # Add new tab

        # Setup tabs
        self.setup_activity_tab()
        self.setup_summary_tab()
        self.setup_graph_tab()  # Setup the new graph tab

        # Create footer with statistics
        self.create_footer()

        self.last_line_count = 0
        self.refresh_interval = 250  # ms, for more responsive polling
        self.load_log()
        self.load_summary()
        self.update_recording_button()
        self.root.after(self.refresh_interval, self.refresh_data)

        # Start the main loop
        self.root.mainloop()

    def setup_activity_tab(self):
        """Setup the Activity Log tab"""
        # Main frame for treeview and scrollbars
        main_frame = tk.Frame(self.activity_frame)
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(main_frame, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        # Vertical scrollbar for activity tree
        self.scroll_y = ttk.Scrollbar(
            main_frame, orient="vertical", command=self.tree.yview)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=self.scroll_y.set)

        # Horizontal scrollbar for activity tree
        self.scroll_x = ttk.Scrollbar(
            self.activity_frame, orient="horizontal", command=self.tree.xview)
        self.scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=self.scroll_x.set)

    def setup_summary_tab(self):
        """Setup the Activity Summary tab"""
        # Main frame for summary treeview and scrollbars
        summary_main_frame = tk.Frame(self.summary_frame)
        summary_main_frame.pack(fill=tk.BOTH, expand=True)

        # Summary info label
        self.summary_info_label = tk.Label(
            self.summary_frame, 
            text="Activity Summary - Categories sorted by total duration (change with right click)",
            font=('Arial', 10, 'bold'),
            pady=5
        )
        self.summary_info_label.pack(side=tk.TOP, fill=tk.X)

        self.summary_tree = ttk.Treeview(summary_main_frame, show="headings")
        self.summary_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        # Bind right-click event to summary tree
        self.summary_tree.bind("<Button-3>", self.on_summary_right_click)

        self.summary_scroll_y = ttk.Scrollbar(
            summary_main_frame, orient="vertical", command=self.summary_tree.yview)
        self.summary_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.summary_tree.configure(yscrollcommand=self.summary_scroll_y.set)

        self.summary_scroll_x = ttk.Scrollbar(
            self.summary_frame, orient="horizontal", command=self.summary_tree.xview)
        self.summary_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.summary_tree.configure(xscrollcommand=self.summary_scroll_x.set)

    def create_footer(self):
        """Create footer with statistics and control buttons"""
        footer_frame = tk.Frame(self.root, relief=tk.SUNKEN, bd=1, height=50)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        footer_frame.pack_propagate(False)  # Prevent frame from shrinking
        
        # Left side - statistics labels
        stats_frame = tk.Frame(footer_frame)
        stats_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Create labels for statistics - make them more visible
        self.first_start_label = tk.Label(
            stats_frame, text="Started: --", anchor='w', 
            font=('Arial', 9), fg='blue')
        self.first_start_label.pack(side=tk.LEFT, padx=(0, 15))

        self.last_stop_label = tk.Label(
            stats_frame, text="Last: --", anchor='w',
            font=('Arial', 9), fg='green')
        self.last_stop_label.pack(side=tk.LEFT, padx=(0, 15))

        self.total_duration_label = tk.Label(
            stats_frame, text="Logged: --", anchor='w',
            font=('Arial', 9), fg='purple')
        self.total_duration_label.pack(side=tk.LEFT, padx=(0, 15))

        self.time_span_label = tk.Label(
            stats_frame, text="Session: --", anchor='w',
            font=('Arial', 9), fg='orange')
        self.time_span_label.pack(side=tk.LEFT, padx=(0, 15))

        self.idle_time_label = tk.Label(
            stats_frame, text="Idle: --", anchor='w',
            font=('Arial', 9), fg='red')
        self.idle_time_label.pack(side=tk.LEFT, padx=(0, 15))

        # Right side - control buttons
        buttons_frame = tk.Frame(footer_frame)
        buttons_frame.pack(side=tk.RIGHT, padx=10, pady=5)

        # Recording status button
        self.recording_btn = tk.Button(
            buttons_frame, 
            text="Logging", 
            command=self.toggle_recording,
            width=12,
            height=1,
            font=('Arial', 9)
        )
        self.recording_btn.pack(side=tk.LEFT, padx=5)

        # Go to folder button
        self.folder_btn = tk.Button(
            buttons_frame, 
            text="Go To Folder", 
            command=self.open_folder,
            width=12,
            height=1,
            font=('Arial', 9)
        )
        self.folder_btn.pack(side=tk.LEFT, padx=5)

    def on_summary_right_click(self, event):
        """Handle right-click on summary tree rows"""
        # Get the item that was clicked
        item = self.summary_tree.identify_row(event.y)
        if not item:
            return
        
        # Get the values of the clicked row
        values = self.summary_tree.item(item, 'values')
        if not values or len(values) < 2:
            return
        
        key = values[0]  # First column is Key
        current_category = values[1]  # Second column is Category
        
        # Get all unique categories from the summary data
        categories = set()
        for child in self.summary_tree.get_children():
            child_values = self.summary_tree.item(child, 'values')
            if len(child_values) >= 2:
                categories.add(child_values[1])
        
        # Sort categories alphabetically
        sorted_categories = sorted(list(categories))
        
        # Create category selection window
        CategorySelector(
            self.root, event, key, current_category, sorted_categories, 
            self.change_category
        )

    def change_category(self, key, new_category):
        """Change the category for a specific key in ActivitySummary.csv and historical data"""
        import __main__
        
        try:
            # Update the category in the logger's app_categories
            if hasattr(__main__, 'logger_instance'):
                logger = __main__.logger_instance
                
                # Store old category for comparison
                old_category = logger.app_categories.get(key, "Unknown")
                
                if old_category == new_category:
                    print(f"Category for '{key}' is already '{new_category}'")
                    return
                
                print(f"Changing category for key '{key}' from '{old_category}' to '{new_category}'")
                
                # Check what keys are actually in the CSV file
                if os.path.exists(logger.log_path):
                    try:
                        with open(logger.log_path, 'r', encoding='utf-8') as f:
                            reader = list(csv.reader(f))
                            if len(reader) > 1:
                                headers = reader[0]
                                if 'ProcessName' in headers:
                                    app_key_index = headers.index('ProcessName')
                                    unique_keys = set()
                                    for row in reader[1:]:
                                        if len(row) > app_key_index:
                                            unique_keys.add(row[app_key_index])
                                
                                print(f"Keys in CSV file: {sorted(unique_keys)}")
                                
                                # Find potential matches
                                potential_matches = []
                                for csv_key in unique_keys:
                                    if hasattr(logger, 'normalize_key_for_matching'):
                                        if logger.normalize_key_for_matching(key, csv_key):
                                            potential_matches.append(csv_key)
                                    elif key == csv_key or (key in csv_key) or (csv_key in key):
                                        potential_matches.append(csv_key)
                                
                                print(f"Potential matches for '{key}': {potential_matches}")
                    except Exception as file_error:
                        print(f"Error reading CSV file: {file_error}")
            
            # Update the category in memory
            logger.app_categories[key] = new_category
            
            # Update historical data in ActivityLog.csv
            updated_count = 0
            if hasattr(logger, 'update_historical_categories'):
                updated_count = logger.update_historical_categories(key, new_category)
            
            # Save the updated categories
            logger.save_app_categories()
            
            # Force immediate refresh of this viewer
            self.refresh_after_category_change()
            
            if updated_count > 0:
                print(f"Successfully updated {updated_count} historical entries")
                # Show success message
                import tkinter.messagebox as messagebox
                messagebox.showinfo("Category Updated", 
                    f"Successfully updated category for '{key}' to '{new_category}'\n"
                    f"Updated {updated_count} historical entries")
            else:
                print(f"No historical entries found for key '{key}'")
                # Show info message
                import tkinter.messagebox as messagebox
                messagebox.showinfo("Category Updated", 
                    f"Category for '{key}' set to '{new_category}'\n"
                    f"No existing historical entries found to update")
                
        except Exception as e:
            print(f"Error changing category: {e}")
            import traceback
            traceback.print_exc()
            import tkinter.messagebox as messagebox
            messagebox.showerror("Error", f"Failed to change category: {e}")

    def refresh_after_category_change(self):
        """Refresh views after a category change"""
        try:
            self.load_log()  # Refresh activity log
            self.load_summary()  # Refresh summary
            self.setup_graph_tab()  # Refresh graph tab
            print("Views refreshed after category change")
        except Exception as e:
            print(f"Error refreshing views after category change: {e}")

    def setup_graph_tab(self):
        """Setup the stacked bar graph tab using the correct log CSV via get_log_path(), skipping 'Inactive' and reducing flicker.
        Only refresh/redraw if window size changed.
        Y axis is in hours. Legend sorted by total duration, with hours shown. Only top 20 processes in legend.
        """
        # Store previous size to avoid unnecessary redraws
        if not hasattr(self, "_graph_prev_size"):
            self._graph_prev_size = (self.graph_frame.winfo_width(), self.graph_frame.winfo_height())

        current_size = (self.graph_frame.winfo_width(), self.graph_frame.winfo_height())
        if hasattr(self, "_graph_canvas") and self._graph_prev_size == current_size:
            # No size change, skip redraw to reduce flicker
            return
        self._graph_prev_size = current_size

        # Clear frame
        for widget in self.graph_frame.winfo_children():
            widget.destroy()

        log_csv_path = self.get_log_path() if hasattr(self, "get_log_path") else self.log_path

        if not os.path.exists(log_csv_path):
            label = tk.Label(self.graph_frame, text=f"{os.path.basename(log_csv_path)} not found", font=('Arial', 12))
            label.pack(padx=10, pady=10)
            return

        try:
            with open(log_csv_path, "r", encoding="utf-8") as f:
                reader = list(csv.reader(f))
                if not reader or len(reader) < 2:
                    label = tk.Label(self.graph_frame, text=f"No data in {os.path.basename(log_csv_path)}", font=('Arial', 12))
                    label.pack(padx=10, pady=10)
                    return
                headers = reader[0]
                rows = reader[1:]

            required_cols = ["Category", "ProcessName", "DurationSeconds"]
            if not all(col in headers for col in required_cols):
                label = tk.Label(self.graph_frame, text=f"Required columns not found in {os.path.basename(log_csv_path)}", font=('Arial', 12))
                label.pack(padx=10, pady=10)
                return

            df = pd.DataFrame(rows, columns=headers)
            df["DurationSeconds"] = pd.to_numeric(df["DurationSeconds"], errors="coerce").fillna(0).astype(int)
            # Filter out 'Inactive' category
            df = df[df["Category"].str.lower() != "inactive"]

            # Group by Category and ProcessName (process name), sum durations
            pivot = df.pivot_table(index="Category", columns="ProcessName", values="DurationSeconds", aggfunc="sum", fill_value=0)
            # Sort categories by total duration descending
            pivot["Total"] = pivot.sum(axis=1)
            pivot = pivot.sort_values("Total", ascending=False)
            pivot = pivot.drop(columns=["Total"])

            if pivot.empty:
                label = tk.Label(self.graph_frame, text="No activity data to display.", font=('Arial', 12))
                label.pack(padx=10, pady=10)
                return

            # Convert seconds to hours for Y axis
            pivot = pivot / 3600.0

            # Sort legend labels by total duration and append hours, only top 20
            process_totals = pivot.sum(axis=0)
            sorted_processs = process_totals.sort_values(ascending=False)
            top_processes = sorted_processs.head(20)
            legend_labels = [
                f"{proc} ({hours:.2f}h)" for proc, hours in top_processes.items()
            ]
            # Reorder columns in pivot to match top_processes, drop others
            pivot = pivot[top_processes.index]

            fig, ax = plt.subplots(figsize=(8, 5), dpi=100)
            bars = pivot.plot(kind="bar", stacked=True, ax=ax)
            ax.set_ylabel("Total Time (hours)")
            ax.set_xlabel("Category")
            ax.set_title("Activity by Category (Stacked by Process)")
            ax.legend(legend_labels, title="Process", bbox_to_anchor=(1.05, 1), loc="upper left")
            plt.tight_layout()

            # Embed in tkinter with reference to avoid flicker on same size
            canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=1)
            self._graph_canvas = canvas
        except Exception as e:
            label = tk.Label(self.graph_frame, text=f"Error loading graph: {e}", font=('Arial', 12))
            label.pack(padx=10, pady=10)

    def on_activity_heading_click(self, col):
        """Handle clicking on activity log column headers for sorting"""
        try:
            # Toggle sort direction if same column, otherwise start with ascending
            if self.sort_column == col:
                self.sort_reverse = not self.sort_reverse
            else:
                self.sort_column = col
                self.sort_reverse = False
            
            # Get all data
            children = list(self.tree.get_children())
            if not children:
                return
            
            # Get column index
            columns = self.tree["columns"]
            if col not in columns:
                return
            col_index = columns.index(col)
            
            # Create list of (values, item_id) for sorting
            data = []
            for child in children:
                values = self.tree.item(child, 'values')
                if len(values) > col_index:
                    sort_value = values[col_index]
                    
                    # Try to convert to appropriate type for sorting
                    if col in ['DurationSeconds', 'Count']:
                        try:
                            sort_value = int(sort_value)
                        except:
                            pass
                    elif col in ['StartTime', 'StopTime']:
                        try:
                            sort_value = datetime.datetime.strptime(sort_value, '%Y-%m-%d %H:%M:%S')
                        except:
                            pass
                    
                    data.append((sort_value, child, values))
            
            # Sort the data
            data.sort(key=lambda x: x[0], reverse=self.sort_reverse)
            
            # Clear and repopulate tree
            self.tree.delete(*children)
            for sort_value, child_id, values in data:
                self.tree.insert("", "end", values=values)
            
            # Update column header to show sort direction
            for column in columns:
                if column == col:
                    direction = " ↓" if self.sort_reverse else " ↑"
                    self.tree.heading(column, text=column + direction)
                else:
                    self.tree.heading(column, text=column)
                    
        except Exception as e:
            print(f"Error sorting column {col}: {e}")

    def load_summary(self):
        """Load ActivitySummary.csv data into the summary tab"""
        if not os.path.exists(self.summary_path):
            # Clear the summary tree if file doesn't exist
            self.summary_tree["columns"] = []
            self.summary_tree.delete(*self.summary_tree.get_children())
            self.summary_info_label.config(text="ActivitySummary.csv not found")
            return

        try:
            with open(self.summary_path, "r", encoding="utf-8") as f:
                reader = list(csv.reader(f))
                if not reader:
                    self.summary_info_label.config(text="ActivitySummary.csv is empty")
                    return

                # Skip comment lines and find headers
                data_start_idx = 0
                headers = []
                metadata = []
                
                for i, row in enumerate(reader):
                    if row and row[0].startswith('#'):
                        # Store metadata for display
                        if len(row) >= 2:
                            metadata.append(f"{row[0]} {row[1]}")
                        else:
                            metadata.append(row[0])
                        continue
                    elif row and not row[0].startswith('#') and row[0]:  # Non-empty, non-comment row
                        headers = row
                        data_start_idx = i + 1
                        break

                if not headers:
                    self.summary_info_label.config(text="No valid headers found in ActivitySummary.csv")
                    return

                # Get data rows
                rows = reader[data_start_idx:] if data_start_idx < len(reader) else []

                # Update info label with metadata
                info_text = "Activity Summary - Categories sorted by total duration (change with right click)"
                if metadata:
                    info_text += " | " + " | ".join(metadata)
                self.summary_info_label.config(text=info_text)

                # Set up columns
                self.summary_tree["columns"] = headers
                for col in headers:
                    self.summary_tree.heading(col, text=col)
                    # Set column widths based on content
                    if col == "Key":
                        self.summary_tree.column(col, width=120, anchor="w")
                    elif col == "Category":
                        self.summary_tree.column(col, width=150, anchor="w")
                    elif col == "Count":
                        self.summary_tree.column(col, width=80, anchor="center")
                    elif col == "Duration":
                        self.summary_tree.column(col, width=120, anchor="center")
                    else:
                        self.summary_tree.column(col, width=100, anchor="w")

                # Remove all old rows
                self.summary_tree.delete(*self.summary_tree.get_children())

                # Insert data rows (already sorted by duration in the file)
                for row in rows:
                    if row and any(cell.strip() for cell in row):  # Skip empty rows
                        # Pad row with empty strings if it's shorter than headers
                        padded_row = row + [''] * (len(headers) - len(row))
                        self.summary_tree.insert("", "end", values=padded_row[:len(headers)])

        except Exception as e:
            print(f"Error loading ActivitySummary.csv: {e}")
            self.summary_info_label.config(text=f"Error loading ActivitySummary.csv: {e}")

    def load_log(self):
        """Load activity log data"""
        if not os.path.exists(self.log_path):
            for col in self.tree["columns"]:
                self.tree.heading(col, text="")
            self.tree.delete(*self.tree.get_children())
            self.update_statistics([])
            self._last_log_mtime = None
            self._last_log_data = None
            return

        try:
            # --- Remember selection before clearing ---
            selected = self.tree.selection()
            selected_key = None
            if selected:
                selected_values = self.tree.item(selected[0], 'values')
                # Use a tuple of the first two columns as a unique key (adjust as needed)
                selected_key = tuple(selected_values[:2]) if selected_values else None

            # Check file modification time
            mtime = os.path.getmtime(self.log_path)
            if hasattr(self, "_last_log_mtime") and self._last_log_mtime == mtime and self._last_log_data is not None:
                headers, rows = self._last_log_data
            else:
                with open(self.log_path, "r", encoding="utf-8") as f:
                    reader = list(csv.reader(f))
                    if not reader:
                        self.update_statistics([])
                        self._last_log_mtime = mtime
                        self._last_log_data = ([], [])
                        return
                    headers = reader[0]
                    rows = reader[1:]
                self._last_log_mtime = mtime
                self._last_log_data = (headers, rows)

            # Reverse the rows so newest is on top
            rows = rows[::-1]

            # Filter out WindowDetails column from display (but keep in CSV)
            display_headers = [h for h in headers if h != "WindowDetails"]
            windowdetails_index = headers.index(
                "WindowDetails") if "WindowDetails" in headers else -1

            # Set up columns (without WindowDetails)
            self.tree["columns"] = display_headers
            for col in display_headers:
                self.tree.heading(col, text=col, 
                                command=lambda c=col: self.on_activity_heading_click(c))
                self.tree.column(col, width=150, anchor="w")

            # Remove all old rows
            self.tree.delete(*self.tree.get_children())

            # Insert new rows (most recent first) - excluding WindowDetails column
            item_id_to_select = None
            for row in rows:
                if windowdetails_index >= 0 and len(row) > windowdetails_index:
                    display_row = row[:windowdetails_index] + row[windowdetails_index + 1:]
                else:
                    display_row = row
                item_id = self.tree.insert("", "end", values=display_row)
                # --- Restore selection if key matches ---
                if selected_key and tuple(display_row[:2]) == selected_key:
                    item_id_to_select = item_id

            self.last_line_count = len(rows)

            # Update statistics (using full rows with WindowDetails for calculations)
            self.update_statistics(rows)

            # --- Restore selection and focus ---
            if item_id_to_select:
                self.tree.selection_set(item_id_to_select)
                self.tree.focus(item_id_to_select)
                self.tree.see(item_id_to_select)

            # Always call after_idle to ensure focus is set after all UI updates
            self.tree.after_idle(lambda: self.tree.focus_set())

        except Exception as e:
            print(f"Error loading log: {e}")

    def update_statistics(self, rows):
        """Update footer statistics based on log data since application start"""
        if not rows:
            self.first_start_label.config(text="Started: --")
            self.last_stop_label.config(text="Last: --")
            self.total_duration_label.config(text="Logged: --")
            self.time_span_label.config(text="Session: --")
            self.idle_time_label.config(text="Idle: --")
            return

        try:
            # Get the logger instance to find app start time
            import __main__
            if hasattr(__main__, 'logger_instance'):
                logger = __main__.logger_instance
                app_start_time = logger.app_start_time
            else:
                # Fallback to first entry if logger not available
                chronological_rows = rows[::-1]
                app_start_time = datetime.datetime.strptime(
                    chronological_rows[0][0], "%Y-%m-%d %H:%M:%S")

            # Filter rows to only include entries since app start
            session_rows = []

            for row in rows:
                try:
                    row_start_time = datetime.datetime.strptime(
                        row[0], "%Y-%m-%d %H:%M:%S")
                    if row_start_time >= app_start_time:
                        session_rows.append(row)
                except:
                    continue

            if not session_rows:
                # No entries since app started
                current_time = datetime.datetime.now()
                session_duration = (
                    current_time - app_start_time).total_seconds()

                self.first_start_label.config(
                    text=f"Started: {app_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
                self.last_stop_label.config(text="Last: --")
                self.total_duration_label.config(text="Logged: 00 00:00:00")
                self.time_span_label.config(
                    text=f"Session: {format_duration(session_duration)}")
                self.idle_time_label.config(
                    text=f"Idle: {format_duration(session_duration)}")
                return

            # Get session statistics
            chronological_session_rows = session_rows[::-1]  # Reverse to chronological order

            # Last activity end time
            last_stop = datetime.datetime.strptime(
                chronological_session_rows[-1][1], "%Y-%m-%d %H:%M:%S")

            # Calculate total logged duration since app start (sum of DurationSeconds)
            total_logged_seconds = sum(
                int(row[2]) for row in session_rows if len(row) > 2 and row[2].isdigit())

            # Calculate session time (time since app started)
            current_time = datetime.datetime.now()
            session_duration = (current_time - app_start_time).total_seconds()

            # Calculate idle time (session time - logged time)
            idle_time_seconds = session_duration - total_logged_seconds

            # Update labels - all times in dd hh:mm:ss format
            self.first_start_label.config(
                text=f"Started: {app_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
            self.last_stop_label.config(
                text=f"Last: {last_stop.strftime('%Y-%m-%d %H:%M:%S')}")
            self.total_duration_label.config(
                text=f"Logged: {format_duration(total_logged_seconds)}")
            self.time_span_label.config(
                text=f"Session: {format_duration(session_duration)}")
            self.idle_time_label.config(
                text=f"Idle: {format_duration(idle_time_seconds)}")

        except Exception as e:
            print(f"Error updating statistics: {e}")
            # Fallback to basic display
            self.first_start_label.config(text="Started: Active")
            self.last_stop_label.config(text="Last: Active")
            self.total_duration_label.config(text=f"Logged: {len(rows)} entries")
            self.time_span_label.config(text="Session: Active")
            self.idle_time_label.config(text="Idle: Calculating...")

    def open_folder(self):
        """Open the folder containing the log file"""
        try:
            folder_path = os.path.dirname(self.log_path)
            os.startfile(folder_path)
        except Exception as e:
            print(f"Error opening folder: {e}")

    def toggle_recording(self):
        """Toggle logging on/off"""
        try:
            import __main__
            if hasattr(__main__, 'logger_instance'):
                logger = __main__.logger_instance
                if logger.running:
                    logger.stop()
                else:
                    logger.start()
            self.update_recording_button()
        except Exception as e:
            print(f"Error toggling recording: {e}")

    def update_recording_button(self):
        """Update recording button text and color based on logging status"""
        try:
            import __main__
            if hasattr(__main__, 'logger_instance'):
                logger = __main__.logger_instance
                if logger.running:
                    self.recording_btn.config(text="Logging", bg='red', fg='white')
                else:
                    self.recording_btn.config(
                        text="Start Logging", bg='blue', fg='white')
        except Exception as e:
            print(f"Error updating recording button: {e}")

    def refresh_data(self):
        """Refresh both Activity Log and Summary data"""
        try:
            # Check if we need to refresh activity log
            if os.path.exists(self.log_path):
                with open(self.log_path, "r", encoding="utf-8") as f:
                    reader = list(csv.reader(f))
                    rows = reader[1:] if len(reader) > 1 else []
                    if len(rows) != self.last_line_count:
                        self.load_log()

            # Refresh summary
            self.load_summary()
            # Refresh graph
            self.setup_graph_tab()

            # Update recording button status
            self.update_recording_button()
            
            # Schedule next refresh
            self.root.after(self.refresh_interval, self.refresh_data)
            
        except Exception as e:
            print(f"Error in refresh_data: {e}")
            # Still schedule next refresh even if there's an error
            self.root.after(self.refresh_interval, self.refresh_data)

    def on_close(self):
        """Handle window close event"""
        # Remove from instances when closed
        if not self.is_duplicate and self.log_path in LogViewer._instances:
            del LogViewer._instances[self.log_path]
        if self.root and self.root.winfo_exists():
            self.root.quit()  # Explicitly stop the event loop
            self.root.destroy() # Then destroy the window and its widgets
