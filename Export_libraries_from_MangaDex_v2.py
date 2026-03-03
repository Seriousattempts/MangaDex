import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
import requests
import pandas as pd
import json
import threading
import time
import os
from datetime import datetime, timedelta
from collections import defaultdict


class MangaDexExtractor:
    def __init__(self):
        self.access_token = None
        self.refresh_token = None
        self.client_id = None
        self.client_secret = None
        self.token_expires_at = None
        self.save_path = None
        self.list_id_entries = []  # Tracks dynamic MDList ID entry widgets

        self.root = tk.Tk()
        self.root.title("MangaDex Data Extractor")
        self.root.geometry("920x950")

        self.setup_ui()

    def setup_ui(self):
        # Scrollable outer frame
        outer = tk.Frame(self.root)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer)
        scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        )

        # ── Credentials ──────────────────────────────────────────────────────
        login_frame = tk.Frame(self.scrollable_frame)
        login_frame.pack(pady=10)

        tk.Label(login_frame, text="MangaDex Credentials",
                 font=("Arial", 14, "bold")).pack(pady=5)

        cred_fields = [
            ("Client ID:", "client_id_entry", ""),
            ("Client Secret:", "client_secret_entry", "*"),
            ("Username:", "username_entry", ""),
            ("Password:", "password_entry", "*"),
        ]
        for label_text, attr_name, show_char in cred_fields:
            frame = tk.Frame(login_frame)
            frame.pack(pady=2)
            tk.Label(frame, text=label_text).pack()
            row = tk.Frame(frame)
            row.pack()
            kw = {"width": 45}
            if show_char:
                kw["show"] = show_char
            entry = tk.Entry(row, **kw)
            entry.pack(side=tk.LEFT)
            setattr(self, attr_name, entry)
            tk.Button(
                row, text="Paste",
                command=lambda e=entry: self.paste_to_entry(e)
            ).pack(side=tk.LEFT, padx=(5, 0))

        # ── Save Location ─────────────────────────────────────────────────────
        save_frame = tk.Frame(self.scrollable_frame)
        save_frame.pack(pady=10)
        tk.Label(save_frame, text="Save Location",
                 font=("Arial", 12, "bold")).pack()
        path_frame = tk.Frame(save_frame)
        path_frame.pack(pady=5)
        self.save_path_var = tk.StringVar(value=os.getcwd())
        tk.Label(path_frame, text="Excel files will be saved to:").pack()
        path_row = tk.Frame(path_frame)
        path_row.pack()
        tk.Label(path_row, textvariable=self.save_path_var,
                 relief="sunken", width=60, anchor="w").pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(path_row, text="Browse",
                  command=self.choose_save_location).pack(side=tk.LEFT)

        # ── Button Sections ───────────────────────────────────────────────────
        button_section = tk.Frame(self.scrollable_frame)
        button_section.pack(pady=10, fill="x", padx=10)

        # Main Extractions
        main_frame = tk.LabelFrame(button_section, text="Main Extractions",
                                   font=("Arial", 10, "bold"))
        main_frame.pack(fill="x", pady=5)
        main_row = tk.Frame(main_frame)
        main_row.pack(pady=5)
        self.all_status_button = tk.Button(
            main_row, text="Extract All Library Status",
            command=self.start_status_extraction,
            bg="blue", fg="white", width=22
        )
        self.all_status_button.pack(side=tk.LEFT, padx=5)
        tk.Button(main_row, text="Clear Log",
                  command=self.clear_log, width=10).pack(side=tk.LEFT, padx=5)

        # Individual Status Extractions
        status_frame = tk.LabelFrame(button_section,
                                     text="Individual Status Extractions",
                                     font=("Arial", 10, "bold"))
        status_frame.pack(fill="x", pady=5)

        row1 = tk.Frame(status_frame)
        row1.pack(pady=2)
        self.reading_button = tk.Button(
            row1, text="Extract Reading",
            command=lambda: self.start_individual_status_extraction("reading"),
            bg="orange", fg="white", width=15)
        self.reading_button.pack(side=tk.LEFT, padx=2)

        self.completed_button = tk.Button(
            row1, text="Extract Completed",
            command=lambda: self.start_individual_status_extraction("completed"),
            bg="purple", fg="white", width=15)
        self.completed_button.pack(side=tk.LEFT, padx=2)

        self.on_hold_button = tk.Button(
            row1, text="Extract On-Hold",
            command=lambda: self.start_individual_status_extraction("on_hold"),
            bg="yellow", fg="black", width=15)
        self.on_hold_button.pack(side=tk.LEFT, padx=2)

        row2 = tk.Frame(status_frame)
        row2.pack(pady=2)
        self.dropped_button = tk.Button(
            row2, text="Extract Dropped",
            command=lambda: self.start_individual_status_extraction("dropped"),
            bg="red", fg="white", width=15)
        self.dropped_button.pack(side=tk.LEFT, padx=2)

        self.plan_to_read_button = tk.Button(
            row2, text="Extract Plan to Read",
            command=lambda: self.start_individual_status_extraction("plan_to_read"),
            bg="cyan", fg="black", width=15)
        self.plan_to_read_button.pack(side=tk.LEFT, padx=2)

        self.re_reading_button = tk.Button(
            row2, text="Extract Re-reading",
            command=lambda: self.start_individual_status_extraction("re_reading"),
            bg="magenta", fg="white", width=15)
        self.re_reading_button.pack(side=tk.LEFT, padx=2)

        # ── Custom MDList Extractions ─────────────────────────────────────────
        mdlist_outer = tk.LabelFrame(button_section,
                                     text="Custom MDList Extractions",
                                     font=("Arial", 10, "bold"))
        mdlist_outer.pack(fill="x", pady=5)

        tk.Label(
            mdlist_outer,
            text="Enter MDList IDs below — each list will be exported to its own Excel file.",
            font=("Arial", 9)
        ).pack(pady=(5, 2))

        self.list_entries_container = tk.Frame(mdlist_outer)
        self.list_entries_container.pack(fill="x", padx=10, pady=5)

        # Seed with one blank entry
        self.add_list_id_entry()

        mdlist_btn_row = tk.Frame(mdlist_outer)
        mdlist_btn_row.pack(pady=5)
        tk.Button(
            mdlist_btn_row, text="+ Add List ID",
            command=self.add_list_id_entry,
            bg="darkgreen", fg="white", width=15
        ).pack(side=tk.LEFT, padx=5)

        self.extract_mdlists_button = tk.Button(
            mdlist_btn_row, text="Extract Custom MDLists",
            command=self.start_mdlist_extraction,
            bg="darkblue", fg="white", width=22
        )
        self.extract_mdlists_button.pack(side=tk.LEFT, padx=5)

        # Master list of all toggleable buttons
        self.all_buttons = [
            self.all_status_button, self.reading_button,
            self.completed_button, self.on_hold_button,
            self.dropped_button, self.plan_to_read_button,
            self.re_reading_button, self.extract_mdlists_button
        ]

        # ── Progress Log ──────────────────────────────────────────────────────
        tk.Label(self.scrollable_frame, text="Progress Log:",
                 font=("Arial", 12, "bold")).pack(pady=(20, 5))
        self.text_widget = scrolledtext.ScrolledText(
            self.scrollable_frame, height=12, state="disabled")
        self.text_widget.pack(fill="both", expand=True, padx=10, pady=5)

    # ── MDList Dynamic Entry Management ──────────────────────────────────────

    def add_list_id_entry(self):
        """Add a new List ID input row to the MDList section."""
        row_frame = tk.Frame(self.list_entries_container)
        row_frame.pack(fill="x", pady=2, anchor="w")

        idx = len(self.list_id_entries) + 1
        label = tk.Label(row_frame, text=f"List ID {idx}:", width=10, anchor="e")
        label.pack(side=tk.LEFT)

        entry = tk.Entry(row_frame, width=42)
        entry.pack(side=tk.LEFT, padx=5)

        tk.Button(row_frame, text="Paste",
                  command=lambda e=entry: self.paste_to_entry(e),
                  width=6).pack(side=tk.LEFT, padx=2)

        tk.Button(row_frame, text="✕", fg="red", width=3,
                  command=lambda rf=row_frame, e=entry: self.remove_list_id_entry(rf, e)
                  ).pack(side=tk.LEFT, padx=2)

        self.list_id_entries.append(entry)

    def remove_list_id_entry(self, row_frame, entry):
        """Remove a List ID row; at least one must remain."""
        if len(self.list_id_entries) <= 1:
            messagebox.showwarning("Cannot Remove",
                                   "At least one List ID field must remain.")
            return
        self.list_id_entries.remove(entry)
        row_frame.destroy()
        # Renumber remaining labels
        for i, row in enumerate(self.list_entries_container.winfo_children(), 1):
            for widget in row.winfo_children():
                if isinstance(widget, tk.Label):
                    widget.config(text=f"List ID {i}:")
                    break

    # ── Generic UI Helpers ────────────────────────────────────────────────────

    def choose_save_location(self):
        folder = filedialog.askdirectory(
            title="Choose folder to save Excel files",
            initialdir=self.save_path_var.get()
        )
        if folder:
            self.save_path_var.set(folder)
            self.print_to_gui(f"✓ Save location set to: {folder}")

    def paste_to_entry(self, entry_widget):
        try:
            content = self.root.clipboard_get()
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, content)
        except tk.TclError:
            messagebox.showwarning("Paste Error",
                                   "No content in clipboard or clipboard is empty.")
        except Exception as e:
            messagebox.showerror("Paste Error", f"Error pasting: {str(e)}")

    def print_to_gui(self, *args, sep=" ", end="\n"):
        text = sep.join(map(str, args)) + end
        print(text.strip())
        self.text_widget.config(state="normal")
        self.text_widget.insert(
            "end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}")
        self.text_widget.see("end")
        self.text_widget.config(state="disabled")
        self.root.update()

    def clear_log(self):
        self.text_widget.config(state="normal")
        self.text_widget.delete(1.0, "end")
        self.text_widget.config(state="disabled")

    def disable_buttons(self):
        for btn in self.all_buttons:
            btn.config(state="disabled")
        self.all_status_button.config(text="Processing...")
        self.reading_button.config(text="Processing...")
        self.completed_button.config(text="Processing...")
        self.on_hold_button.config(text="Processing...")
        self.dropped_button.config(text="Processing...")
        self.plan_to_read_button.config(text="Processing...")
        self.re_reading_button.config(text="Processing...")
        self.extract_mdlists_button.config(text="Processing...")

    def enable_buttons(self):
        for btn in self.all_buttons:
            btn.config(state="normal")
        self.all_status_button.config(text="Extract All Library Status")
        self.reading_button.config(text="Extract Reading")
        self.completed_button.config(text="Extract Completed")
        self.on_hold_button.config(text="Extract On-Hold")
        self.dropped_button.config(text="Extract Dropped")
        self.plan_to_read_button.config(text="Extract Plan to Read")
        self.re_reading_button.config(text="Extract Re-reading")
        self.extract_mdlists_button.config(text="Extract Custom MDLists")

    # ── Authentication ────────────────────────────────────────────────────────

    def authenticate(self):
        self.print_to_gui("Starting authentication...")
        self.client_id = self.client_id_entry.get().strip()
        self.client_secret = self.client_secret_entry.get().strip()
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not all([self.client_id, self.client_secret, username, password]):
            self.print_to_gui("ERROR: All credential fields are required!")
            return False

        creds = {
            "grant_type": "password",
            "username": username,
            "password": password,
            "client_id": self.client_id,
            "client_secret": self.client_secret
        }
        try:
            self.print_to_gui("Sending authentication request...")
            r = requests.post(
                "https://auth.mangadex.org/realms/mangadex/protocol/openid-connect/token",
                data=creds
            )
            if r.status_code == 200:
                rj = r.json()
                self.access_token = rj["access_token"]
                self.refresh_token = rj["refresh_token"]
                self.token_expires_at = datetime.now() + timedelta(minutes=14)
                self.print_to_gui("✓ Authentication successful!")
                self.print_to_gui(
                    f"✓ Token refreshes at: {self.token_expires_at.strftime('%H:%M:%S')}")
                return True
            else:
                self.print_to_gui(
                    f"✗ Authentication failed: {r.status_code} - {r.text}")
                return False
        except Exception as e:
            self.print_to_gui(f"✗ Authentication error: {str(e)}")
            return False

    def refresh_access_token(self):
        self.print_to_gui("🔄 Refreshing access token...")
        creds = {
            "grant_type": "refresh_token",
            "refresh_token": self.refresh_token,
            "client_id": self.client_id,
            "client_secret": self.client_secret
        }
        try:
            r = requests.post(
                "https://auth.mangadex.org/realms/mangadex/protocol/openid-connect/token",
                data=creds
            )
            if r.status_code == 200:
                rj = r.json()
                self.access_token = rj["access_token"]
                if "refresh_token" in rj:
                    self.refresh_token = rj["refresh_token"]
                self.token_expires_at = datetime.now() + timedelta(minutes=14)
                self.print_to_gui("✓ Token refreshed!")
                return True
            else:
                self.print_to_gui(
                    f"✗ Token refresh failed: {r.status_code} - {r.text}")
                return False
        except Exception as e:
            self.print_to_gui(f"✗ Token refresh error: {str(e)}")
            return False

    def check_and_refresh_token(self):
        if self.token_expires_at is None:
            return False
        if datetime.now() >= self.token_expires_at:
            self.print_to_gui("⏰ Token expired, refreshing...")
            return self.refresh_access_token()
        return True

    def make_authenticated_request(self, url, **kwargs):
        if not self.check_and_refresh_token():
            raise Exception("Failed to refresh token")
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {self.access_token}"
        return requests.get(url, headers=headers, **kwargs)

    # ── Feature 1: Latest Read Chapter & Feature 2: Scanlation Group ─────────

    def fetch_all_read_chapters_grouped(self, manga_ids):
        """
        GET /manga/read?grouped=true&ids[]=...
        Returns dict: manga_id -> [chapter_id, ...], or None on format error.
        """
        self.print_to_gui(
            f"Fetching read chapter data for {len(manga_ids)} manga (grouped)...")
        grouped_data = {}

        for i in range(0, len(manga_ids), 100):
            batch = manga_ids[i:i + 100]
            # requests encodes list-of-tuples as repeated query params: ids[]=a&ids[]=b
            params = [("ids[]", mid) for mid in batch] + [("grouped", "true")]
            try:
                if not self.check_and_refresh_token():
                    break
                resp = self.make_authenticated_request(
                    "https://api.mangadex.org/manga/read", params=params)
                if resp.status_code == 200:
                    data = resp.json().get("data", {})
                    if isinstance(data, dict):
                        grouped_data.update(data)
                        self.print_to_gui(
                            f"  Batch {i // 100 + 1}: {len(data)} manga have read chapters")
                    else:
                        # API returned flat list — grouped param unsupported
                        self.print_to_gui(
                            "⚠️ Grouped response not available, switching to per-manga fallback...")
                        return None
                else:
                    self.print_to_gui(
                        f"⚠️ Read batch {i // 100 + 1} failed: {resp.status_code}")
            except Exception as e:
                self.print_to_gui(f"⚠️ Error in grouped read fetch: {str(e)}")
            time.sleep(0.5)

        self.print_to_gui(
            f"  {len(grouped_data)} manga have at least one read chapter recorded")
        return grouped_data

    def fetch_read_chapters_for_manga(self, manga_id):
        """
        Fallback: GET /manga/{id}/read
        Returns a flat list of chapter IDs for a single manga.
        """
        try:
            resp = self.make_authenticated_request(
                f"https://api.mangadex.org/manga/{manga_id}/read")
            if resp.status_code == 200:
                return resp.json().get("data", [])
        except Exception as e:
            self.print_to_gui(
                f"⚠️ Error fetching read chapters for {manga_id}: {str(e)}")
        return []

    def fetch_chapter_details_batch(self, chapter_ids):
        """
        GET /chapter?ids[]=...&limit=100
        Fetches chapter attributes (chapter number, volume) and relationships
        (scanlation_group) for up to 100 chapter IDs per call.
        """
        if not chapter_ids:
            return []
        unique_ids = list(set(chapter_ids))
        total_batches = (len(unique_ids) + 99) // 100
        self.print_to_gui(
            f"Fetching details for {len(unique_ids)} chapters "
            f"({total_batches} batch{'es' if total_batches > 1 else ''})...")
        chapters = []

        for i in range(0, len(unique_ids), 100):
            batch = unique_ids[i:i + 100]
            params = [("ids[]", cid) for cid in batch] + [("limit", "100")]
            try:
                if not self.check_and_refresh_token():
                    break
                resp = self.make_authenticated_request(
                    "https://api.mangadex.org/chapter", params=params)
                if resp.status_code == 200:
                    chapters.extend(resp.json().get("data", []))
                else:
                    self.print_to_gui(
                        f"⚠️ Chapter batch {i // 100 + 1} failed: {resp.status_code}")
            except Exception as e:
                self.print_to_gui(f"⚠️ Chapter batch error: {str(e)}")
            time.sleep(0.3)

        self.print_to_gui(f"✓ Retrieved {len(chapters)} chapter records")
        return chapters

    def fetch_group_details(self, group_ids):
        """
        GET /group?ids[]=...&limit=100
        Returns dict: group_id -> group_name
        """
        unique_ids = list(set(filter(None, group_ids)))
        if not unique_ids:
            return {}
        self.print_to_gui(
            f"Fetching names for {len(unique_ids)} scanlation groups...")
        groups = {}

        for i in range(0, len(unique_ids), 100):
            batch = unique_ids[i:i + 100]
            params = [("ids[]", gid) for gid in batch] + [("limit", "100")]
            try:
                if not self.check_and_refresh_token():
                    break
                resp = self.make_authenticated_request(
                    "https://api.mangadex.org/group", params=params)
                if resp.status_code == 200:
                    for g in resp.json().get("data", []):
                        gid = g.get("id", "")
                        gname = (g.get("attributes") or {}).get("name", "")
                        if gid:
                            groups[gid] = gname
                else:
                    self.print_to_gui(
                        f"⚠️ Group batch failed: {resp.status_code}")
            except Exception as e:
                self.print_to_gui(f"⚠️ Group batch error: {str(e)}")
            time.sleep(0.3)

        self.print_to_gui(f"✓ Retrieved {len(groups)} scanlation group names")
        return groups

    def build_read_chapter_map(self, manga_ids):
        """
        Builds and returns a dict:
          manga_id -> {
              latestReadChapter, latestReadVolume,
              scanlationGroup, scanlationGroupUrl
          }

        Steps:
          1. Batch-fetch all read chapter IDs grouped by manga via /manga/read?grouped=true
          2. If that fails, fall back to per-manga /manga/{id}/read calls
          3. Batch-fetch chapter details via /chapter?ids[]=...
          4. Batch-fetch scanlation group names via /group?ids[]=...
          5. For each manga, identify the chapter with the highest chapter number
             as the "latest read chapter" and attach its scanlation group
        """
        if not manga_ids:
            return {}
        self.print_to_gui(
            f"\n--- Building read chapter map for {len(manga_ids)} manga ---")

        # Step 1: Grouped fetch
        grouped = self.fetch_all_read_chapters_grouped(manga_ids)

        # Step 2: Fallback to per-manga if needed
        if grouped is None:
            self.print_to_gui("Falling back to per-manga read chapter fetch...")
            grouped = {}
            for mid in manga_ids:
                ch_ids = self.fetch_read_chapters_for_manga(mid)
                if ch_ids:
                    grouped[mid] = ch_ids
                time.sleep(0.1)

        if not grouped:
            self.print_to_gui("No read chapter data found.")
            return {}

        # Collect all unique chapter IDs across all manga
        all_ch_ids = list({cid for ids in grouped.values() for cid in ids})
        if not all_ch_ids:
            return {}

        # Step 3: Fetch chapter details
        chapters = self.fetch_chapter_details_batch(all_ch_ids)

        # Step 4: Build a lookup table and collect group IDs
        chapter_lookup = {}
        group_ids_seen = set()

        for ch in chapters:
            cid = ch.get("id", "")
            if not cid:
                continue
            attrs = ch.get("attributes") or {}
            sg_id = None
            for rel in ch.get("relationships") or []:
                if rel.get("type") == "scanlation_group":
                    sg_id = rel.get("id")
                    if sg_id:
                        group_ids_seen.add(sg_id)
                    break
            chapter_lookup[cid] = {
                "chapter": attrs.get("chapter", ""),
                "volume": attrs.get("volume", ""),
                "scanlation_group_id": sg_id
            }

        # Step 4b: Fetch group names
        group_names = self.fetch_group_details(list(group_ids_seen))

        # Step 5: Per-manga, find the chapter with the highest numeric chapter value
        read_map = {}
        for mid, ch_ids in grouped.items():
            best = None
            best_val = -1.0
            for cid in ch_ids:
                if cid not in chapter_lookup:
                    continue
                cd = chapter_lookup[cid]
                try:
                    val = float(cd["chapter"]) if cd["chapter"] else -1.0
                except (ValueError, TypeError):
                    val = -1.0
                if val > best_val:
                    best_val = val
                    best = cd

            if best:
                sg_id = best.get("scanlation_group_id")
                read_map[mid] = {
                    "latestReadChapter": best.get("chapter", ""),
                    "latestReadVolume": best.get("volume", ""),
                    "scanlationGroup": group_names.get(sg_id, "") if sg_id else "",
                    "scanlationGroupUrl": (
                        f"https://mangadex.org/group/{sg_id}" if sg_id else ""
                    )
                }

        self.print_to_gui(
            f"✓ Read chapter map complete: {len(read_map)} manga mapped\n")
        return read_map

    # ── Feature 3: MDList Fetching ────────────────────────────────────────────

    def fetch_mdlist_manga_ids(self, list_id):
        """
        GET /list/{list_id}
        Extracts manga IDs from the list's relationships array.
        Returns (list_name, [manga_id, ...]) or (None, None) on failure.
        """
        self.print_to_gui(f"Fetching MDList: {list_id}...")
        try:
            if not self.check_and_refresh_token():
                return None, None
            resp = self.make_authenticated_request(
                f"https://api.mangadex.org/list/{list_id}")
            if resp.status_code == 200:
                data = resp.json().get("data", {})
                list_name = (data.get("attributes") or {}).get("name", list_id)
                manga_ids = [
                    rel["id"]
                    for rel in (data.get("relationships") or [])
                    if rel.get("type") == "manga"
                ]
                self.print_to_gui(
                    f"✓ List '{list_name}': {len(manga_ids)} manga found")
                return list_name, manga_ids
            else:
                self.print_to_gui(
                    f"✗ Failed to fetch MDList {list_id}: "
                    f"{resp.status_code} - {resp.text}")
                return None, None
        except Exception as e:
            self.print_to_gui(f"✗ Error fetching MDList {list_id}: {str(e)}")
            return None, None

    # ── Core Manga Fetching ───────────────────────────────────────────────────

    def fetch_manga_statuses(self):
        """GET /manga/status — returns all user library manga grouped by status."""
        self.print_to_gui("Fetching manga statuses for all categories...")
        manga_status_ids = defaultdict(list)
        try:
            if not self.check_and_refresh_token():
                return manga_status_ids
            resp = self.make_authenticated_request(
                "https://api.mangadex.org/manga/status")
            if resp.status_code == 200:
                for manga_id, status in resp.json().get("statuses", {}).items():
                    manga_status_ids[status].append(manga_id)
                for status, ids in manga_status_ids.items():
                    self.print_to_gui(
                        f"✓ Retrieved {len(ids)} manga IDs for status '{status}'")
            else:
                self.print_to_gui(
                    f"✗ Failed to fetch statuses: {resp.status_code} - {resp.text}")
        except Exception as e:
            self.print_to_gui(f"✗ Error fetching statuses: {str(e)}")
        return manga_status_ids

    def fetch_manga_details(self, manga_ids):
        """GET /manga/{id} for each ID — returns list of manga data objects."""
        self.print_to_gui(
            f"Fetching detailed manga info for {len(manga_ids)} manga...")
        manga_details = []
        for i, mid in enumerate(manga_ids, 1):
            try:
                if i % 50 == 0 and not self.check_and_refresh_token():
                    self.print_to_gui("✗ Token refresh failed, stopping detail fetch")
                    break
                resp = self.make_authenticated_request(
                    f"https://api.mangadex.org/manga/{mid}")
                if resp.status_code == 200:
                    manga_details.append(resp.json().get("data", {}))
                    if i % 50 == 0:
                        self.print_to_gui(
                            f"  Fetched {i}/{len(manga_ids)} manga details...")
                else:
                    self.print_to_gui(
                        f"✗ Failed for {mid}: {resp.status_code}")
            except Exception as e:
                self.print_to_gui(f"✗ Error for {mid}: {str(e)}")
            time.sleep(0.1)
        self.print_to_gui(
            f"✓ Fetched details for {len(manga_details)} manga")
        return manga_details

    # ── Data Processing ───────────────────────────────────────────────────────

    def flatten_nested_data(self, data_list, read_chapter_map=None):
        """
        Converts the nested manga JSON objects into flat dicts suitable for
        a DataFrame / Excel sheet.

        If read_chapter_map is provided (manga_id -> read data), the following
        columns are added to each record:
          - latestReadChapter  (highest chapter number the user has read)
          - latestReadVolume   (volume of that chapter)
          - scanlationGroup    (name of the group that translated it)
          - scanlationGroupUrl (link to that group on MangaDex)
        """
        self.print_to_gui("Processing and flattening manga data...")
        flattened = []

        for i, manga in enumerate(data_list, 1):
            try:
                if i % 50 == 0:
                    self.print_to_gui(
                        f"  Processing manga {i}/{len(data_list)}...")

                attrs = manga.get("attributes") or {}
                manga_id_raw = manga.get("id", "")

                flat = {
                    "type": manga.get("type", ""),
                    "id": f"https://mangadex.org/title/{manga_id_raw}"
                }

                # Titles (one column per language variant)
                titles = attrs.get("title") or {}
                if isinstance(titles, dict):
                    for j, (_, v) in enumerate(titles.items(), 1):
                        if v:
                            flat[f"title{j}"] = v

                # Descriptions (one column per language variant)
                descs = attrs.get("description") or {}
                if isinstance(descs, dict):
                    for j, (_, v) in enumerate(descs.items(), 1):
                        if v:
                            flat[f"description{j}"] = v

                # External links (al, mal, mu, etc.)
                links = attrs.get("links") or {}
                if isinstance(links, dict):
                    link_counts = defaultdict(int)
                    for lt, lv in links.items():
                        if lv:
                            col = self.get_link_column_name(lt)
                            link_counts[col] += 1
                            col_name = (col if link_counts[col] == 1
                                        else f"{col}{link_counts[col]}")
                            flat[col_name] = self.create_link_url(lt, lv)

                # Simple metadata fields
                flat["publicationDemographic"] = attrs.get(
                    "publicationDemographic", "")
                flat["status"] = attrs.get("status", "")
                flat["year"] = attrs.get("year", "")
                flat["lastVolume"] = attrs.get("lastVolume", "")
                flat["lastChapter"] = attrs.get("lastChapter", "")

                # ── Latest read chapter + scanlation group (Feature 1 & 2) ──
                if read_chapter_map and manga_id_raw in read_chapter_map:
                    rd = read_chapter_map[manga_id_raw]
                    flat["latestReadChapter"] = rd.get("latestReadChapter", "")
                    flat["latestReadVolume"] = rd.get("latestReadVolume", "")
                    flat["scanlationGroup"] = rd.get("scanlationGroup", "")
                    flat["scanlationGroupUrl"] = rd.get("scanlationGroupUrl", "")
                else:
                    flat["latestReadChapter"] = ""
                    flat["latestReadVolume"] = ""
                    flat["scanlationGroup"] = ""
                    flat["scanlationGroupUrl"] = ""

                # Tags
                tags = attrs.get("tags") or []
                if isinstance(tags, list):
                    for j, tag in enumerate(tags, 1):
                        if tag and isinstance(tag, dict):
                            tag_name = (
                                tag.get("attributes") or {}
                            ).get("name") or {}
                            if isinstance(tag_name, dict):
                                tag_text = next(
                                    (v for v in tag_name.values() if v), "")
                                if tag_text:
                                    flat[f"tags{j}"] = tag_text

                # Author / artist relationships
                rels = manga.get("relationships") or []
                rel_counts = defaultdict(int)
                for rel in rels:
                    if not rel or not isinstance(rel, dict):
                        continue
                    rt = rel.get("type", "")
                    rid = rel.get("id", "")
                    if rt == "cover_art":
                        continue
                    if rt in ("author", "artist") and rid:
                        rel_counts[rt] += 1
                        col = rt if rel_counts[rt] == 1 else f"{rt}{rel_counts[rt]}"
                        flat[col] = f"https://mangadex.org/author/{rid}"

                flattened.append(flat)

            except Exception as e:
                self.print_to_gui(
                    f"✗ Error processing manga "
                    f"{manga.get('id', 'unknown')}: {str(e)}")
                try:
                    flattened.append({
                        "type": manga.get("type", ""),
                        "id": f"https://mangadex.org/title/{manga.get('id', '')}"
                    })
                except Exception:
                    pass

        self.print_to_gui(
            f"✓ Successfully processed {len(flattened)} manga records")
        return flattened

    def export_to_excel_with_status(self, data, label):
        """
        Writes flattened manga records to a timestamped .xlsx file.
        Column order: type/id → titles → descriptions → links →
                      metadata → read chapter data → author/artist → tags
        """
        try:
            self.print_to_gui(f"Creating Excel file for '{label}'...")
            df = pd.DataFrame(data)

            base_cols = ["type", "id"]
            title_cols = sorted(c for c in df.columns if c.startswith("title"))
            desc_cols = sorted(
                c for c in df.columns if c.startswith("description"))
            other_base = [
                "publicationDemographic", "status", "year",
                "lastVolume", "lastChapter"
            ]
            # New columns — latest read progress and scanlation group
            read_cols = [
                "latestReadChapter", "latestReadVolume",
                "scanlationGroup", "scanlationGroupUrl"
            ]
            rel_cols = sorted(
                c for c in df.columns if c.startswith(("author", "artist")))
            tag_cols = sorted(c for c in df.columns if c.startswith("tags"))

            named = set(
                base_cols + title_cols + desc_cols +
                other_base + read_cols + rel_cols + tag_cols
            )
            link_cols = sorted(c for c in df.columns if c not in named)

            ordered = (base_cols + title_cols + desc_cols +
                       link_cols + other_base + read_cols + rel_cols + tag_cols)
            final_cols = [c for c in ordered if c in df.columns]
            df = df[final_cols]

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"mangadex_{label}_series_{ts}.xlsx"
            full_path = os.path.join(self.save_path_var.get(), filename)
            df.to_excel(full_path, index=False, engine="openpyxl")
            return full_path

        except Exception as e:
            self.print_to_gui(
                f"✗ Error exporting to Excel for '{label}': {str(e)}")
            return None

    # ── URL Helpers ───────────────────────────────────────────────────────────

    def get_correct_mangaupdates_url(self, url):
        if (url.startswith("https://www.mangaupdates.com/series/")
                and "id=" not in url):
            return url
        if "id=" in url:
            try:
                self.print_to_gui(f"🔗 Resolving MangaUpdates URL: {url}")
                resp = requests.get(url, allow_redirects=True, timeout=10)
                if resp.status_code == 200:
                    final = resp.url.rstrip("/")
                    parts = final.split("/")
                    if len(parts) > 5:
                        final = "/".join(parts[:5])
                    self.print_to_gui(f"✓ Resolved to: {final}")
                    return final
                else:
                    return url
            except Exception as e:
                self.print_to_gui(f"⚠️ URL resolve error: {str(e)}")
                return url
        return url

    def create_link_url(self, link_type, link_value):
        clean = (link_value.replace("series/", "")
                 if isinstance(link_value, str) else str(link_value))
        if link_type == "mu":
            if clean.isdigit():
                return self.get_correct_mangaupdates_url(
                    f"https://www.mangaupdates.com/series.html?id={clean}")
            return self.get_correct_mangaupdates_url(
                f"https://www.mangaupdates.com/series/{clean}")
        mappings = {
            "al": f"https://anilist.co/manga/{clean}",
            "kt": f"https://kitsu.app/manga/{clean}",
            "ap": f"https://www.anime-planet.com/manga/{clean}",
            "bw": f"https://bookwalker.jp/series/{clean}",
            "mal": f"https://myanimelist.net/manga/{clean}",
            "amz": str(link_value),
            "raw": str(link_value),
            "ebj": str(link_value),
            "cdj": str(link_value)
        }
        return mappings.get(link_type, str(link_value))

    def get_link_column_name(self, link_type):
        return {
            "al": "anilist", "kt": "kitsu", "mu": "mangaupdates",
            "ap": "animeplanet", "bw": "bookwalker", "mal": "myanimelist",
            "amz": "amazon", "raw": "raw", "ebj": "ebookjapan", "cdj": "cdjapan"
        }.get(link_type, link_type)

    # ── Orchestration ─────────────────────────────────────────────────────────

    def process_and_export_status_manga(self, manga_status_ids):
        """Process every status group and export each to its own Excel file."""
        for status, ids in manga_status_ids.items():
            self.print_to_gui(
                f"\n--- Processing status '{status}' ({len(ids)} manga) ---")
            if not ids:
                self.print_to_gui(f"No manga IDs for '{status}', skipping.")
                continue

            manga_details = self.fetch_manga_details(ids)
            if not manga_details:
                self.print_to_gui(
                    f"✗ No manga details for '{status}', skipping.")
                continue

            read_map = self.build_read_chapter_map(ids)
            flattened = self.flatten_nested_data(manga_details, read_map)
            if not flattened:
                self.print_to_gui(
                    f"✗ No flattened data for '{status}', skipping.")
                continue

            filename = self.export_to_excel_with_status(flattened, status)
            if filename:
                self.print_to_gui(
                    f"✓ Exported {len(flattened)} manga for '{status}' → {filename}")
            else:
                self.print_to_gui(f"✗ Failed to export '{status}'")

    def process_and_export_individual_status(self, status, manga_ids):
        """Process a single status group and export to Excel."""
        self.print_to_gui(
            f"Processing status '{status}' with {len(manga_ids)} manga...")

        if not manga_ids:
            self.print_to_gui(f"No manga IDs for '{status}'. Aborted.")
            return False

        manga_details = self.fetch_manga_details(manga_ids)
        if not manga_details:
            self.print_to_gui(
                f"✗ No manga details for '{status}'. Aborted.")
            return False

        read_map = self.build_read_chapter_map(manga_ids)
        flattened = self.flatten_nested_data(manga_details, read_map)
        if not flattened:
            self.print_to_gui(f"✗ No flattened data for '{status}'. Aborted.")
            return False

        filename = self.export_to_excel_with_status(flattened, status)
        if filename:
            self.print_to_gui(
                f"✓ Exported {len(flattened)} manga for '{status}' → {filename}")
            return True
        return False

    # ── Thread Launchers ──────────────────────────────────────────────────────

    def start_status_extraction(self):
        self.disable_buttons()
        threading.Thread(
            target=self.extract_status_data, daemon=True).start()

    def start_individual_status_extraction(self, status):
        self.disable_buttons()
        threading.Thread(
            target=lambda: self.extract_individual_status_data(status),
            daemon=True
        ).start()

    def start_mdlist_extraction(self):
        self.disable_buttons()
        threading.Thread(
            target=self.extract_mdlist_data, daemon=True).start()

    # ── Extraction Entry Points ───────────────────────────────────────────────

    def extract_status_data(self):
        """Full-library extraction: all statuses → one Excel file each."""
        try:
            if not self.authenticate():
                return
            manga_status_ids = self.fetch_manga_statuses()
            if not manga_status_ids:
                self.print_to_gui("✗ No manga status data. Aborted.")
                return
            self.process_and_export_status_manga(manga_status_ids)
            total = sum(1 for ids in manga_status_ids.values() if ids)
            messagebox.showinfo(
                "Success",
                f"All library status extraction complete!\n"
                f"{total} Excel file(s) created."
            )
        except Exception as e:
            self.print_to_gui(f"✗ Unexpected error: {str(e)}")
            messagebox.showerror("Error", str(e))
        finally:
            self.enable_buttons()

    def extract_individual_status_data(self, target_status):
        """Single-status extraction."""
        try:
            if not self.authenticate():
                return
            manga_status_ids = self.fetch_manga_statuses()
            if not manga_status_ids:
                return
            if target_status not in manga_status_ids:
                self.print_to_gui(
                    f"✗ No manga for status '{target_status}'.")
                messagebox.showwarning(
                    "No Data",
                    f"No manga found for status '{target_status}'.")
                return
            manga_ids = manga_status_ids[target_status]
            success = self.process_and_export_individual_status(
                target_status, manga_ids)
            if success:
                messagebox.showinfo(
                    "Success",
                    f"Extracted {len(manga_ids)} manga for '{target_status}'.")
            else:
                messagebox.showerror(
                    "Error",
                    f"Failed to extract '{target_status}'.")
        except Exception as e:
            self.print_to_gui(f"✗ Unexpected error: {str(e)}")
            messagebox.showerror("Error", str(e))
        finally:
            self.enable_buttons()

    def extract_mdlist_data(self):
        """
        MDList extraction:
          For each non-empty List ID entered by the user:
            1. Fetch the list from /list/{id} to get its manga IDs
            2. Fetch full manga details for those IDs
            3. Build read chapter map (latest read chapter + scanlation group)
            4. Flatten and export to a separate Excel file named after the list
        """
        try:
            list_ids = [
                e.get().strip()
                for e in self.list_id_entries
                if e.get().strip()
            ]
            if not list_ids:
                messagebox.showwarning(
                    "No List IDs",
                    "Please enter at least one MDList ID.")
                return

            if not self.authenticate():
                return

            success_count = 0
            for list_id in list_ids:
                self.print_to_gui(f"\n{'=' * 50}")
                self.print_to_gui(f"Processing MDList: {list_id}")
                self.print_to_gui(f"{'=' * 50}")

                list_name, manga_ids = self.fetch_mdlist_manga_ids(list_id)
                if manga_ids is None:
                    self.print_to_gui(
                        f"✗ Skipping list {list_id} due to fetch error.")
                    continue
                if not manga_ids:
                    self.print_to_gui(
                        f"⚠️ MDList '{list_name}' is empty. Skipping.")
                    continue

                manga_details = self.fetch_manga_details(manga_ids)
                if not manga_details:
                    self.print_to_gui(
                        f"✗ No manga details for list '{list_name}'. Skipping.")
                    continue

                read_map = self.build_read_chapter_map(manga_ids)
                flattened = self.flatten_nested_data(manga_details, read_map)
                if not flattened:
                    self.print_to_gui(
                        f"✗ No data to export for list '{list_name}'.")
                    continue

                # Sanitise the list name for use in a filename
                safe_name = "".join(
                    c for c in (list_name or list_id)
                    if c.isalnum() or c in (" ", "_", "-")
                ).strip().replace(" ", "_")

                filename = self.export_to_excel_with_status(
                    flattened, f"mdlist_{safe_name}")
                if filename:
                    self.print_to_gui(
                        f"✓ '{list_name}' — {len(flattened)} manga → {filename}")
                    success_count += 1
                else:
                    self.print_to_gui(
                        f"✗ Export failed for list '{list_name}'")

            messagebox.showinfo(
                "Done",
                f"MDList extraction complete!\n"
                f"{success_count} of {len(list_ids)} list(s) exported successfully."
            )

        except Exception as e:
            self.print_to_gui(f"✗ Unexpected error: {str(e)}")
            messagebox.showerror("Error", str(e))
        finally:
            self.enable_buttons()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    try:
        import requests
        import pandas
        import openpyxl
    except ImportError:
        print("Missing required package. Install with: "
              "pip install requests pandas openpyxl")
        exit(1)

    app = MangaDexExtractor()
    app.run()
