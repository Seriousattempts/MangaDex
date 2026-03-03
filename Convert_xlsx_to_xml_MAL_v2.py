import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import pandas as pd
import os
from datetime import datetime
import re
import gzip


class MangaXMLGenerator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Manga XML Generator")
        self.root.geometry("900x700")

        self.file_entries = []  # Store file entry widgets

        self.setup_ui()

    def setup_ui(self):
        # Main frame with scrollbar
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # User info section
        user_frame = tk.LabelFrame(main_frame, text="User Information", font=("Arial", 12, "bold"))
        user_frame.pack(fill="x", pady=5)

        # User ID with paste button
        user_id_frame = tk.Frame(user_frame)
        user_id_frame.pack(pady=2)
        tk.Label(user_id_frame, text="User ID:").pack(side=tk.LEFT)
        self.user_id_entry = tk.Entry(user_id_frame, width=20)
        self.user_id_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(user_id_frame, text="Paste", command=lambda: self.paste_to_entry(self.user_id_entry)).pack(
            side=tk.LEFT)

        # User Name with paste button
        user_name_frame = tk.Frame(user_frame)
        user_name_frame.pack(pady=2)
        tk.Label(user_name_frame, text="User Name:").pack(side=tk.LEFT)
        self.user_name_entry = tk.Entry(user_name_frame, width=20)
        self.user_name_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(user_name_frame, text="Paste", command=lambda: self.paste_to_entry(self.user_name_entry)).pack(
            side=tk.LEFT)

        # File selection section
        file_frame = tk.LabelFrame(main_frame, text="Excel Files", font=("Arial", 12, "bold"))
        file_frame.pack(fill="both", expand=True, pady=5)

        # Create canvas and scrollbar for scrollable content
        canvas_frame = tk.Frame(file_frame)
        canvas_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(canvas_frame, height=200)
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

        # Save location section
        save_frame = tk.LabelFrame(main_frame, text="Save Location", font=("Arial", 12, "bold"))
        save_frame.pack(fill="x", pady=5)

        save_path_frame = tk.Frame(save_frame)
        save_path_frame.pack(pady=5)

        self.save_path_var = tk.StringVar()
        self.save_path_var.set(os.getcwd())

        tk.Label(save_path_frame, text="XML file will be saved to:").pack()
        path_display_frame = tk.Frame(save_path_frame)
        path_display_frame.pack()

        self.path_label = tk.Label(path_display_frame, textvariable=self.save_path_var,
                                   relief="sunken", width=70, anchor="w")
        self.path_label.pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(path_display_frame, text="Browse", command=self.choose_save_location).pack(side=tk.LEFT)

        # Generate button
        generate_frame = tk.Frame(main_frame)
        generate_frame.pack(fill="x", pady=10)
        tk.Button(generate_frame, text="Generate XML", command=self.generate_xml,
                  bg="green", fg="white", font=("Arial", 14, "bold")).pack()

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

        # Status and Priority section
        options_frame = tk.Frame(path_frame)
        options_frame.pack(fill="x", pady=2)

        # Status dropdown
        status_frame = tk.Frame(options_frame)
        status_frame.pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(status_frame, text="Status:").pack(anchor="w")
        status_var = tk.StringVar()
        status_dropdown = ttk.Combobox(status_frame, textvariable=status_var, width=15,
                                       values=["Reading", "Completed", "On-Hold", "Dropped", "Plan to Read"])
        status_dropdown.pack()
        status_dropdown.set("Reading")  # Default value

        # Priority dropdown
        priority_frame = tk.Frame(options_frame)
        priority_frame.pack(side=tk.LEFT)
        tk.Label(priority_frame, text="Priority:").pack(anchor="w")
        priority_var = tk.StringVar()
        priority_dropdown = ttk.Combobox(priority_frame, textvariable=priority_var, width=10,
                                         values=["Low", "Medium", "High"])
        priority_dropdown.pack()
        priority_dropdown.set("Low")  # Default value

        # Store the entry widgets
        entry_data = {
            'number': entry_num,
            'path_entry': file_path_entry,
            'status_var': status_var,
            'priority_var': priority_var,
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
        except:
            pass  # Silently handle any clipboard errors

    def choose_save_location(self):
        folder_path = filedialog.askdirectory(
            title="Choose folder to save XML file",
            initialdir=self.save_path_var.get()
        )
        if folder_path:
            self.save_path_var.set(folder_path)

    def extract_mal_id(self, mal_url):
        """Extract MyAnimeList ID from URL"""
        if not mal_url or not isinstance(mal_url, str):
            return None

        # Remove trailing slash and split by "/"
        parts = mal_url.rstrip('/').split('/')
        if parts:
            try:
                # Get the last part and try to convert to int
                mal_id = parts[-1]
                int(mal_id)  # Validate it's a number
                return mal_id
            except (ValueError, IndexError):
                return None
        return None

    def clean_volume_chapter(self, value):
        """Remove apostrophes from volume/chapter values"""
        if not value or pd.isna(value):
            return "0"

        # Convert to string and remove apostrophes
        clean_value = str(value).replace("'", "")

        # Try to extract number, default to 0 if not possible
        try:
            # Handle cases where there might be additional text
            numbers = re.findall(r'\d+', clean_value)
            if numbers:
                return numbers[0]
            else:
                return "0"
        except:
            return "0"

    def get_my_times_read(self, status):
        """Return my_times_read value based on status"""
        if status in ["Reading", "Completed", "Dropped"]:
            return "1"
        else:
            return "0"

    def process_excel_files(self):
        """Process all selected Excel files and return manga data"""
        all_manga_data = []
        status_counts = {"Reading": 0, "Completed": 0, "On-Hold": 0, "Dropped": 0, "Plan to Read": 0}

        for entry_data in self.file_entries:
            file_path = entry_data['path_entry'].get().strip()
            if not file_path:
                continue

            try:
                # Read Excel file
                df = pd.read_excel(file_path)

                # Check if required columns exist
                required_columns = ['myanimelist', 'title1', 'latestReadVolume', 'latestReadChapter']
                missing_columns = [col for col in required_columns if col not in df.columns]

                if missing_columns:
                    messagebox.showwarning("Missing Columns",
                                           f"File {os.path.basename(file_path)} is missing columns: {', '.join(missing_columns)}")
                    continue

                # Filter rows that have myanimelist values
                df_filtered = df[df['myanimelist'].notna() & (df['myanimelist'] != '')]

                status = entry_data['status_var'].get()
                priority = entry_data['priority_var'].get()

                # Process each row
                for _, row in df_filtered.iterrows():
                    mal_id = self.extract_mal_id(row['myanimelist'])
                    if mal_id:
                        manga_data = {
                            'mal_id': mal_id,
                            'title': str(row['title1']) if pd.notna(row['title1']) else "",
                            'volumes': self.clean_volume_chapter(row['latestReadVolume']),
                            'chapters': self.clean_volume_chapter(row['latestReadChapter']),
                            'status': status,
                            'priority': priority
                        }
                        all_manga_data.append(manga_data)
                        status_counts[status] += 1

            except Exception as e:
                messagebox.showerror("Error", f"Error processing file {file_path}: {str(e)}")
                continue

        return all_manga_data, status_counts

    def generate_xml(self):
        """Generate the XML file and compressed version"""
        try:
            # Validate inputs
            user_id = self.user_id_entry.get().strip()
            user_name = self.user_name_entry.get().strip()

            if not user_id or not user_name:
                messagebox.showerror("Error", "User ID and User Name are required!")
                return

            # Check if at least one file is selected
            has_files = any(entry['path_entry'].get().strip() for entry in self.file_entries)
            if not has_files:
                messagebox.showerror("Error", "Please select at least one Excel file!")
                return

            # Process Excel files
            manga_data, status_counts = self.process_excel_files()

            if not manga_data:
                messagebox.showerror("Error", "No valid manga data found in the selected files!")
                return

            # Create XML content
            xml_content = self.create_xml_content(user_id, user_name, manga_data, status_counts)

            # Generate filenames with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            xml_filename = f"myanimelist_export_{timestamp}.xml"
            gz_filename = f"myanimelist_export_{timestamp}.xml.gz"
            save_directory = self.save_path_var.get()
            xml_full_path = os.path.join(save_directory, xml_filename)
            gz_full_path = os.path.join(save_directory, gz_filename)

            # Save XML file
            with open(xml_full_path, 'w', encoding='utf-8') as f:
                f.write(xml_content)

            # Create compressed .gz file
            with open(xml_full_path, 'rb') as f_in:
                with gzip.open(gz_full_path, 'wb') as f_out:
                    f_out.write(f_in.read())

            messagebox.showinfo("Success",
                                f"XML files generated successfully!\n\n"
                                f"XML File: {xml_full_path}\n"
                                f"Compressed File: {gz_full_path}\n\n"
                                f"Total manga: {len(manga_data)}\n"
                                f"Reading: {status_counts['Reading']}\n"
                                f"Completed: {status_counts['Completed']}\n"
                                f"On-Hold: {status_counts['On-Hold']}\n"
                                f"Dropped: {status_counts['Dropped']}\n"
                                f"Plan to Read: {status_counts['Plan to Read']}")

        except Exception as e:
            messagebox.showerror("Error", f"Error generating XML: {str(e)}")

    def create_xml_content(self, user_id, user_name, manga_data, status_counts):
        """Create the XML content string matching the exact format"""
        total_manga = len(manga_data)

        xml_lines = [
            '<?xml version="1.0" encoding="UTF-8" ?>',
            '\t',
            '\t<!--',
            '\t Created by XML Export feature at MyAnimeList.net',
            '\t Version 1.1.0',
            '\t-->',
            '\t',
            '\t<myanimelist>',
            '\t',
            '\t\t<myinfo>',
            f'\t\t\t<user_id>{user_id}</user_id>',
            f'\t\t\t<user_name>{user_name}</user_name>',
            f'\t\t\t<user_export_type>2</user_export_type>',
            f'\t\t\t<user_total_manga>{total_manga}</user_total_manga>',
            f'\t\t\t<user_total_reading>{status_counts["Reading"]}</user_total_reading>',
            f'\t\t\t<user_total_completed>{status_counts["Completed"]}</user_total_completed>',
            f'\t\t\t<user_total_onhold>{status_counts["On-Hold"]}</user_total_onhold>',
            f'\t\t\t<user_total_dropped>{status_counts["Dropped"]}</user_total_dropped>',
            f'\t\t\t<user_total_plantoread>{status_counts["Plan to Read"]}</user_total_plantoread>',
            '\t\t</myinfo>',
            '\t'
        ]

        # Add manga entries
        for manga in manga_data:
            # Get the my_times_read value based on status
            times_read = self.get_my_times_read(manga["status"])

            xml_lines.extend([
                '\t',
                '\t\t\t\t<manga>',
                f'\t\t\t\t\t<manga_mangadb_id>{manga["mal_id"]}</manga_mangadb_id>',
                f'\t\t\t\t\t<manga_title><![CDATA[{manga["title"]}]]></manga_title>',
                f'\t\t\t\t\t<my_read_volumes>{manga["volumes"]}</my_read_volumes>',
                f'\t\t\t\t\t<my_read_chapters>{manga["chapters"]}</my_read_chapters>',
                '\t\t\t\t\t<my_start_date>0000-00-00</my_start_date>',
                '\t\t\t\t\t<my_finish_date>0000-00-00</my_finish_date>',
                '\t\t\t\t\t<my_scanalation_group><![CDATA[]]></my_scanalation_group>',
                '\t\t\t\t\t<my_score>0</my_score>',
                '\t\t\t\t\t<my_storage></my_storage>',
                '\t\t\t\t\t<my_retail_volumes>0</my_retail_volumes>',
                f'\t\t\t\t\t<my_status>{manga["status"]}</my_status>',
                '\t\t\t\t\t<my_comments><![CDATA[]]></my_comments>',
                f'\t\t\t\t\t<my_times_read>{times_read}</my_times_read>',
                '\t\t\t\t\t<my_tags><![CDATA[]]></my_tags>',
                f'\t\t\t\t\t<my_priority>{manga["priority"]}</my_priority>',
                '\t\t\t\t\t<my_reread_value></my_reread_value>',
                '\t\t\t\t\t<my_rereading>NO</my_rereading>',
                '\t\t\t\t\t<my_discuss>YES</my_discuss>',
                '\t\t\t\t\t<my_sns>default</my_sns>',
                '\t\t\t\t\t<update_on_import>1</update_on_import>',
                '\t\t\t\t</manga>',
                '\t\t\t'
            ])

        xml_lines.extend([
            '\t\t</myanimelist>'
        ])

        return '\n'.join(xml_lines)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    # Required packages
    required_packages = ["pandas", "openpyxl"]

    try:
        import pandas
        import openpyxl
    except ImportError as e:
        print(f"Missing required package. Install with: pip install {' '.join(required_packages)}")
        exit(1)

    app = MangaXMLGenerator()
    app.run()
