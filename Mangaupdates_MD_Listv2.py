import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, ttk
import requests
import pandas as pd
import json
import threading
import time
import os
import re
from datetime import datetime
from collections import defaultdict


class MangaUpdatesListManager:
    def __init__(self):
        self.session_token = None
        self.session_cookie = None
        self.refresh_cookie = None
        self.base_url = "https://api.mangaupdates.com/v1"
        self.file_entries = []  # Store file entry widgets

        # Map list names to their IDs
        self.list_mapping = {
            "Reading List": 0,
            "Wish List": 1,
            "Complete List": 2,
            "Unfinished List": 3,
            "On Hold List": 4
        }

        # Create main window
        self.root = tk.Tk()
        self.root.title("MangaUpdates List Manager")
        self.root.geometry("800x700")

        self.setup_ui()

    def setup_ui(self):
        # Main frame
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Login form
        login_frame = tk.LabelFrame(main_frame, text="MangaUpdates Credentials", font=("Arial", 12, "bold"))
        login_frame.pack(fill="x", pady=5)

        # Username with paste button
        username_frame = tk.Frame(login_frame)
        username_frame.pack(pady=2)
        tk.Label(username_frame, text="Username:").pack(side=tk.LEFT)
        self.username_entry = tk.Entry(username_frame, width=30)
        self.username_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(username_frame, text="Paste", command=lambda: self.paste_to_entry(self.username_entry)).pack(
            side=tk.LEFT)

        # Password with paste button
        password_frame = tk.Frame(login_frame)
        password_frame.pack(pady=2)
        tk.Label(password_frame, text="Password:").pack(side=tk.LEFT)
        self.password_entry = tk.Entry(password_frame, width=30, show="*")
        self.password_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(password_frame, text="Paste", command=lambda: self.paste_to_entry(self.password_entry)).pack(
            side=tk.LEFT)

        # File selection section
        file_frame = tk.LabelFrame(main_frame, text="Excel Files for List Updates", font=("Arial", 12, "bold"))
        file_frame.pack(fill="both", expand=True, pady=5)

        # Create canvas and scrollbar for scrollable content
        canvas_frame = tk.Frame(file_frame)
        canvas_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(canvas_frame, height=250)
        self.scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Add file button
        add_button_frame = tk.Frame(file_frame)
        add_button_frame.pack(fill="x", pady=5)
        tk.Button(add_button_frame, text="Add Excel File", command=self.add_file_entry,
                  bg="blue", fg="white").pack()

        # Add first file entry
        self.add_file_entry()

        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)

        self.update_button = tk.Button(button_frame, text="Update Lists from Excel",
                                       command=self.start_list_update, bg="green", fg="white",
                                       font=("Arial", 12, "bold"))
        self.update_button.pack(side=tk.LEFT, padx=5)

        tk.Button(button_frame, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)

        # Output text area
        tk.Label(main_frame, text="Progress Log:", font=("Arial", 12, "bold")).pack(pady=(20, 5))
        self.text_widget = scrolledtext.ScrolledText(main_frame, height=15, state="disabled")
        self.text_widget.pack(fill="both", expand=True, padx=10, pady=5)

    def add_file_entry(self):
        entry_num = len(self.file_entries) + 1

        # Create frame for this file entry
        file_entry_frame = tk.Frame(self.scrollable_frame, relief="ridge", bd=1)
        file_entry_frame.pack(fill="x", padx=5, pady=3)

        # Number label
        number_label = tk.Label(file_entry_frame, text=f"{entry_num}",
                                width=3, font=("Arial", 10, "bold"), bg="lightblue")
        number_label.pack(side=tk.LEFT, padx=2)

        # File path section
        path_frame = tk.Frame(file_entry_frame)
        path_frame.pack(side=tk.LEFT, fill="x", expand=True, padx=5)

        tk.Label(path_frame, text="File:").pack(anchor="w")
        path_sub_frame = tk.Frame(path_frame)
        path_sub_frame.pack(fill="x")

        file_path_entry = tk.Entry(path_sub_frame, width=50)
        file_path_entry.pack(side=tk.LEFT, fill="x", expand=True)

        browse_btn = tk.Button(path_sub_frame, text="Browse",
                               command=lambda: self.browse_file(file_path_entry))
        browse_btn.pack(side=tk.LEFT, padx=(5, 0))

        # List selection dropdown
        list_frame = tk.Frame(path_frame)
        list_frame.pack(fill="x", pady=2)

        tk.Label(list_frame, text="Target List:").pack(anchor="w")
        list_var = tk.StringVar()
        list_dropdown = ttk.Combobox(list_frame, textvariable=list_var, width=20,
                                     values=["Reading List", "Wish List", "Complete List",
                                             "Unfinished List", "On Hold List"])
        list_dropdown.pack()
        list_dropdown.set("Reading List")  # Default value

        # Store the entry widgets
        entry_data = {
            'number': entry_num,
            'path_entry': file_path_entry,
            'list_var': list_var,
            'frame': file_entry_frame,
            'number_label': number_label
        }

        self.file_entries.append(entry_data)

        # Update canvas scroll region
        self.root.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def browse_file(self, entry_widget):
        file_path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)

    def paste_to_entry(self, entry_widget):
        """Paste clipboard content to the specified entry widget"""
        try:
            clipboard_content = self.root.clipboard_get()
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, clipboard_content)
            self.print_to_gui(f"✓ Pasted content to field")
        except tk.TclError:
            messagebox.showwarning("Paste Error", "No content in clipboard or clipboard is empty")
        except Exception as e:
            messagebox.showerror("Paste Error", f"Error pasting from clipboard: {str(e)}")

    def print_to_gui(self, *args, sep=" ", end="\n"):
        """Display text in the GUI text widget similar to print function"""
        text = sep.join(map(str, args)) + end
        print(text.strip())  # Also print to console
        self.text_widget.config(state="normal")
        self.text_widget.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}")
        self.text_widget.see("end")
        self.text_widget.config(state="disabled")
        self.root.update()

    def clear_log(self):
        self.text_widget.config(state="normal")
        self.text_widget.delete(1.0, "end")
        self.text_widget.config(state="disabled")

    def start_list_update(self):
        """Start the list update process in a separate thread"""
        self.update_button.config(state="disabled", text="Updating...")
        thread = threading.Thread(target=self.update_lists_from_excel)
        thread.daemon = True
        thread.start()

    def clean_numeric_value(self, value):
        """Clean apostrophes and extract numeric value from strings like '343"""
        if not value or pd.isna(value):
            return 0
        try:
            # Convert to string, remove apostrophes and extract digits
            s = str(value).replace("'", "")
            match = re.findall(r"\d+", s)
            if match:
                return int(match[0])  # Take first number found
            else:
                return 0
        except:
            return 0

    def authenticate(self):
        """Authenticate with MangaUpdates API using PUT method"""
        self.print_to_gui("Starting authentication process...")

        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not username or not password:
            self.print_to_gui("ERROR: Username and password are required!")
            return False

        # MangaUpdates API authentication using PUT method
        auth_data = {
            "username": username,
            "password": password
        }

        try:
            self.print_to_gui("Sending authentication request...")
            response = requests.put(f"{self.base_url}/account/login", json=auth_data)

            if response.status_code == 200:
                auth_response = response.json()
                self.print_to_gui(f"✓ Authentication response received")

                # Extract authentication details
                context = auth_response.get("context", {})
                self.session_token = context.get("session_token")

                # Extract cookies from response headers
                if 'Set-Cookie' in response.headers:
                    cookies = response.headers['Set-Cookie']
                    # Parse cookies to extract session and refresh cookies
                    for cookie in cookies.split(','):
                        if 'session=' in cookie:
                            self.session_cookie = cookie.split('session=')[1].split(';')[0]
                        elif 'refresh=' in cookie:
                            self.refresh_cookie = cookie.split('refresh=')[1].split(';')[0]

                if self.session_token:
                    self.print_to_gui("✓ Authentication successful!")
                    self.print_to_gui(f"✓ Session token acquired")
                    if self.refresh_cookie:
                        self.print_to_gui(f"✓ Refresh cookie acquired")
                    return True
                else:
                    self.print_to_gui("✗ Authentication failed: No session token received")
                    return False
            else:
                self.print_to_gui(f"✗ Authentication failed: {response.status_code} - {response.text}")
                return False

        except Exception as e:
            self.print_to_gui(f"✗ Authentication error: {str(e)}")
            return False

    def refresh_session(self):
        """Refresh session using refresh cookie"""
        if not self.refresh_cookie:
            self.print_to_gui("✗ No refresh cookie available")
            return False

        self.print_to_gui("🔄 Refreshing session...")

        try:
            headers = {
                "Cookie": f"refresh={self.refresh_cookie}",
                "Content-Type": "application/json"
            }

            response = requests.put(f"{self.base_url}/account/refresh", headers=headers)

            if response.status_code == 200:
                auth_response = response.json()
                context = auth_response.get("context", {})
                self.session_token = context.get("session_token")

                # Update cookies if provided
                if 'Set-Cookie' in response.headers:
                    cookies = response.headers['Set-Cookie']
                    for cookie in cookies.split(','):
                        if 'session=' in cookie:
                            self.session_cookie = cookie.split('session=')[1].split(';')[0]
                        elif 'refresh=' in cookie:
                            self.refresh_cookie = cookie.split('refresh=')[1].split(';')[0]

                if self.session_token:
                    self.print_to_gui("✓ Session refreshed successfully!")
                    return True
                else:
                    self.print_to_gui("✗ Session refresh failed: No session token received")
                    return False
            else:
                self.print_to_gui(f"✗ Session refresh failed: {response.status_code} - {response.text}")
                return False

        except Exception as e:
            self.print_to_gui(f"✗ Session refresh error: {str(e)}")
            return False

    def make_authenticated_request(self, endpoint, data=None, params=None, method="GET"):
        """Make an authenticated request to MangaUpdates API with auto-refresh"""
        headers = {
            "Content-Type": "application/json"
        }

        # Add authentication
        if self.session_token:
            headers["Authorization"] = f"Bearer {self.session_token}"

        if self.session_cookie:
            headers["Cookie"] = f"session={self.session_cookie}"

        url = f"{self.base_url}/{endpoint}"

        # Make the request
        if method.upper() == "GET":
            response = requests.get(url, headers=headers, params=params)
        elif method.upper() == "POST":
            response = requests.post(url, headers=headers, json=data, params=params)
        elif method.upper() == "PUT":
            response = requests.put(url, headers=headers, json=data, params=params)
        elif method.upper() == "PATCH":
            response = requests.patch(url, headers=headers, json=data, params=params)
        else:
            response = requests.request(method, url, headers=headers, params=params)

        # If unauthorized, try to refresh session
        if response.status_code == 401 and self.refresh_cookie:
            self.print_to_gui("🔄 Session expired, attempting refresh...")
            if self.refresh_session():
                # Retry the request with new session
                if self.session_token:
                    headers["Authorization"] = f"Bearer {self.session_token}"
                if self.session_cookie:
                    headers["Cookie"] = f"session={self.session_cookie}"

                if method.upper() == "GET":
                    response = requests.get(url, headers=headers, params=params)
                elif method.upper() == "POST":
                    response = requests.post(url, headers=headers, json=data, params=params)
                elif method.upper() == "PUT":
                    response = requests.put(url, headers=headers, json=data, params=params)
                elif method.upper() == "PATCH":
                    response = requests.patch(url, headers=headers, json=data, params=params)
                else:
                    response = requests.request(method, url, headers=headers, params=params)

        return response

    def extract_series_info_from_url(self, url, current_index, total_urls, filename):
        """Extract series identifier and name from MangaUpdates URL with enhanced progress tracking"""
        try:
            self.print_to_gui(f"🔍 [{filename}] ({current_index}/{total_urls}) Extracting info from: {url}")
            response = requests.get(url, timeout=10)

            if response.status_code == 200:
                content = response.text

                # Look for the JSON pattern in the page source
                # Pattern: "identifier":20518774890,"name":"Houketsu no Ronny"
                identifier_pattern = r'"identifier":(\d+),"name":"([^"]+)"'
                match = re.search(identifier_pattern, content)

                if match:
                    identifier = match.group(1)
                    name = match.group(2)
                    self.print_to_gui(f"✓ [{filename}] Found - ID: {identifier}, Name: {name}")
                    return {"series_id": int(identifier), "title": name}
                else:
                    self.print_to_gui(f"⚠️ [{filename}] Could not find identifier/name pattern in page source")
                    return None
            else:
                self.print_to_gui(f"⚠️ [{filename}] Failed to fetch page (status {response.status_code})")
                return None

        except Exception as e:
            self.print_to_gui(f"⚠️ [{filename}] Error extracting info from URL: {str(e)}")
            return None

    def process_excel_files_for_update(self):
        """Process Excel files to extract series for list updates with volume/chapter data"""
        all_series_data = []  # List of all series objects formatted for API
        file_summaries = {}  # {filename: {list_name: count}}

        for entry_data in self.file_entries:
            file_path = entry_data['path_entry'].get().strip()
            list_selection = entry_data['list_var'].get().strip()

            if not file_path or not list_selection:
                continue

            if list_selection not in self.list_mapping:
                self.print_to_gui(f"⚠️ Unknown list selection: {list_selection}")
                continue

            list_id = self.list_mapping[list_selection]
            filename = os.path.basename(file_path)
            file_summaries[filename] = {}

            try:
                self.print_to_gui(f"📖 Processing file: {filename} for {list_selection}")

                # Read Excel file
                df = pd.read_excel(file_path)

                # Check if required columns exist
                required_columns = ['mangaupdates']
                missing_columns = [col for col in required_columns if col not in df.columns]

                if missing_columns:
                    self.print_to_gui(f"⚠️ File {filename} is missing columns: {', '.join(missing_columns)}")
                    continue

                # Filter rows that have mangaupdates values
                df_filtered = df[df['mangaupdates'].notna() & (df['mangaupdates'] != '')]
                total_urls = len(df_filtered)

                if total_urls == 0:
                    self.print_to_gui(f"⚠️ File {filename} has no valid MangaUpdates URLs")
                    continue

                self.print_to_gui(f"📊 Found {total_urls} MangaUpdates URLs in {filename}")

                successful_extractions = 0

                # Process each row with mangaupdates URL
                for index, (_, row) in enumerate(df_filtered.iterrows(), 1):
                    mu_url = str(row['mangaupdates']).strip()
                    if mu_url and mu_url != 'nan':
                        series_info = self.extract_series_info_from_url(mu_url, index, total_urls, filename)
                        if series_info:
                            # Extract volume and chapter data from Excel
                            last_volume = self.clean_numeric_value(row.get('latestReadVolume', 0))
                            last_chapter = self.clean_numeric_value(row.get('latestReadChapter', 0))

                            # Format series data for API
                            series_data = {
                                "series": {
                                    "id": series_info["series_id"],
                                    "title": series_info["title"]
                                },
                                "list_id": list_id,
                                "status": {
                                    "volume": last_volume,
                                    "chapter": last_chapter,
                                    "increment_volume": 0,
                                    "increment_chapter": 0
                                },
                                "priority": 0
                            }
                            all_series_data.append(series_data)
                            successful_extractions += 1

                        time.sleep(0.5)  # Rate limiting for web scraping

                file_summaries[filename][list_selection] = successful_extractions
                self.print_to_gui(
                    f"✅ Completed {filename}: {successful_extractions}/{total_urls} series successfully extracted for {list_selection}")

            except Exception as e:
                self.print_to_gui(f"✗ Error processing file {filename}: {str(e)}")
                continue

        return all_series_data, file_summaries

    def add_series_to_lists_batch(self, all_series_data):
        """Add series to lists in batches of 25 with proper rate limiting"""
        if not all_series_data:
            self.print_to_gui("⚠️ No series data to add")
            return 0

        total_series = len(all_series_data)
        batch_size = 25

        # Split into batches of 25
        batches = [all_series_data[i:i + batch_size] for i in range(0, len(all_series_data), batch_size)]

        self.print_to_gui(
            f"📝 Adding {total_series} series to lists in {len(batches)} batches of up to {batch_size} series each...")

        total_successful = 0

        for batch_num, batch in enumerate(batches, 1):
            try:
                self.print_to_gui(f"📦 Processing batch {batch_num}/{len(batches)} ({len(batch)} series)...")

                # Show sample data for first batch
                if batch_num == 1:
                    sample_data = batch[:2]
                    self.print_to_gui(f"🔍 Sample batch data:")
                    for i, sample in enumerate(sample_data, 1):
                        self.print_to_gui(
                            f"  Sample {i}: {sample['series']['title']} (Vol: {sample['status']['volume']}, Ch: {sample['status']['chapter']})")

                # Send batch to API
                response = self.make_authenticated_request("lists/series", data=batch, method="POST")

                if response.status_code in [200, 201]:
                    batch_success = len(batch)
                    total_successful += batch_success
                    self.print_to_gui(f"  ✅ Batch {batch_num} successful: {batch_success} series added")

                    # Parse response for detailed results if available
                    try:
                        response_data = response.json()
                        if isinstance(response_data, list):
                            actual_success = len([r for r in response_data if r.get('success', True)])
                            if actual_success != batch_success:
                                self.print_to_gui(f"  ⚠️ API reported {actual_success}/{batch_success} successful")
                                total_successful = total_successful - batch_success + actual_success
                    except:
                        pass

                else:
                    self.print_to_gui(f"  ✗ Batch {batch_num} failed: {response.status_code} - {response.text}")

                # Rate limiting - wait between batches
                if batch_num < len(batches):
                    self.print_to_gui(f"  ⏳ Waiting 2 seconds before next batch...")
                    time.sleep(2)

            except Exception as e:
                self.print_to_gui(f"  ✗ Error processing batch {batch_num}: {str(e)}")
                continue

        self.print_to_gui(f"✅ Batch processing complete: {total_successful}/{total_series} series successfully added")
        return total_successful

    def update_lists_from_excel(self):
        """Main process to update lists from Excel files"""
        try:
            # Step 1: Authenticate
            if not self.authenticate():
                return

            # Step 2: Process Excel files
            all_series_data, file_summaries = self.process_excel_files_for_update()

            if not all_series_data:
                self.print_to_gui("✗ No valid data found in Excel files for list updates")
                return

            # Step 3: Add series using batch method
            total_series = len(all_series_data)
            self.print_to_gui(f"🚀 Preparing to add {total_series} series to lists")

            successful_additions = self.add_series_to_lists_batch(all_series_data)

            # Step 4: Show detailed summary
            self.print_to_gui("=" * 60)
            self.print_to_gui("📊 FINAL SUMMARY")
            self.print_to_gui("=" * 60)

            # Summary by file
            self.print_to_gui("📁 By File:")
            for filename, lists in file_summaries.items():
                for list_name, count in lists.items():
                    self.print_to_gui(f"  {filename} → {list_name}: {count} series extracted")

            self.print_to_gui("")

            # Summary by list
            self.print_to_gui("📋 By List:")
            list_counts = {}
            for series_data in all_series_data:
                list_id = series_data["list_id"]
                list_name = next((name for name, id in self.list_mapping.items() if id == list_id), f"List {list_id}")
                list_counts[list_name] = list_counts.get(list_name, 0) + 1

            for list_name, count in list_counts.items():
                self.print_to_gui(f"  {list_name}: {count} series attempted")

            self.print_to_gui("")
            self.print_to_gui(
                f"🎯 OVERALL TOTAL: {successful_additions}/{total_series} series successfully added to lists")

            # Show completion message
            messagebox.showinfo("Update Complete",
                                f"List update process completed!\n\n"
                                f"Successfully added {successful_additions}/{total_series} series\n"
                                f"Processed {len(file_summaries)} Excel files\n"
                                f"Updated {len(list_counts)} lists")

        except Exception as e:
            self.print_to_gui(f"✗ Unexpected error during list update: {str(e)}")
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")

        finally:
            self.update_button.config(state="normal", text="Update Lists from Excel")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    # Required packages
    required_packages = ["requests", "pandas", "openpyxl"]

    try:
        import requests
        import pandas
        import openpyxl
    except ImportError as e:
        print(f"Missing required package. Install with: pip install {' '.join(required_packages)}")
        exit(1)

    app = MangaUpdatesListManager()
    app.run()
