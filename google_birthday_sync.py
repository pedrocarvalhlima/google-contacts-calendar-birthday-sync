import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
import pickle
import threading
import platform
from rapidfuzz import process, fuzz
from ics import Calendar

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

CSV_PATH = "./calendar.csv"
SCOPES = ["https://www.googleapis.com/auth/contacts"]
LANGUAGES = ["en", "pt", "es"]  # Add more as needed


def authenticate_google():
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return creds


def get_contacts(service):
    results = (
        service.people()
        .connections()
        .list(resourceName="people/me", pageSize=500, personFields="names,birthdays")
        .execute()
    )
    return results.get("connections", [])


def get_contact_details(service, resource_name):
    # Fetch full contact details including etag
    return (
        service.people()
        .get(resourceName=resource_name, personFields="names,birthdays,metadata")
        .execute()
    )


def update_birthday(service, contact, birthday_str):
    year, month, day = birthday_str.split("-")
    # Only include month and day, omit year
    birthday = {
        "etag": contact.get("etag")
        or (contact.get("metadata", {}).get("sources", [{}])[0].get("etag")),
        "birthdays": [{"date": {"month": int(month), "day": int(day)}}],
    }
    service.people().updateContact(
        resourceName=contact["resourceName"], updatePersonFields="birthdays", body=birthday
    ).execute()


def create_contact(service, name, birthday_str):
    year, month, day = birthday_str.split("-")
    # Only include month and day, omit year
    contact_body = {
        "names": [{"displayName": name}],
        "birthdays": [{"date": {"month": int(month), "day": int(day)}}],
    }
    service.people().createContact(body=contact_body).execute()


class AutocompleteEntry(tk.Entry):
    def __init__(self, contacts, textvariable, parent, callback, *args, **kwargs):
        super().__init__(parent, textvariable=textvariable, *args, **kwargs)
        self.contacts = contacts
        self.callback = callback
        self.listbox = None
        self.textvariable = textvariable
        self.parent = parent
        self.bind("<KeyRelease>", self.on_keyrelease)
        self.bind("<FocusOut>", self.hide_listbox)
        # Bind both Control-BackSpace and Command-BackSpace for cross-platform word deletion
        self.bind("<Control-BackSpace>", self.ctrl_backspace)
        self.bind("<Command-BackSpace>", self.ctrl_backspace)
        self.bind("<Button-4>", self.on_mousewheel)  # Linux scroll up
        self.bind("<Button-5>", self.on_mousewheel)  # Linux scroll down
        self.bind("<MouseWheel>", self.on_mousewheel)  # Windows/Mac scroll

    def ctrl_backspace(self, event):
        # Remove word before cursor
        idx = self.index(tk.INSERT)
        text = self.get()
        left = text[:idx]
        right = text[idx:]
        left = left.rstrip()
        if not left:
            self.delete(0, tk.END)
            return "break"
        # Find last space
        space_idx = left.rfind(" ")
        new_left = left[: space_idx + 1] if space_idx != -1 else ""
        self.delete(0, tk.END)
        self.insert(0, new_left + right)
        self.icursor(len(new_left))
        return "break"

    def on_mousewheel(self, event):
        if self.listbox:
            if platform.system() == "Darwin":  # macOS
                if event.num == 4 or (hasattr(event, "delta") and event.delta > 0):
                    self.listbox.yview_scroll(-1, "units")
                elif event.num == 5 or (hasattr(event, "delta") and event.delta < 0):
                    self.listbox.yview_scroll(1, "units")
            elif platform.system() == "Windows":
                self.listbox.yview_scroll(-1 * (event.delta // 120), "units")
            else:  # Linux
                if event.num == 4:
                    self.listbox.yview_scroll(-1, "units")
                elif event.num == 5:
                    self.listbox.yview_scroll(1, "units")

    def on_keyrelease(self, event=None):
        value = self.textvariable.get()
        if value == "":
            self.hide_listbox()
            return
        matches = [c for c in self.contacts if value.lower() in c.lower()]
        if not matches:
            self.hide_listbox()
            return
        if not self.listbox:
            self.listbox = tk.Listbox(self.winfo_toplevel(), height=5)
            self.listbox.bind("<<ListboxSelect>>", self.on_select)
        self.listbox.delete(0, tk.END)
        for match in matches:
            self.listbox.insert(tk.END, match)
        # Place listbox below entry
        x = self.winfo_rootx() - self.winfo_toplevel().winfo_rootx()
        y = self.winfo_rooty() - self.winfo_toplevel().winfo_rooty() + self.winfo_height()
        self.listbox.place(x=x, y=y)
        self.listbox.lift()
        self.listbox.activate(0)

    def on_select(self, event):
        if self.listbox.curselection():
            value = self.listbox.get(self.listbox.curselection()[0])
            self.textvariable.set(value)
            self.callback(value)
            self.hide_listbox()

    def hide_listbox(self, event=None):
        if self.listbox:
            self.listbox.place_forget()


class CalendarSyncApp(tk.Tk):
    def __init__(self, df, contacts, service):
        super().__init__()
        self.title("Calendar Birthday Sync")
        self.geometry("1000x700")
        self.minsize(800, 500)
        self.df = df
        self.contacts = contacts
        self.service = service
        self.contact_names = [c["names"][0]["displayName"] for c in contacts if c.get("names")]
        self.entries = []
        self.processed_entries = []
        self.errors = []
        self.progress = 0
        self.total = 0
        self.removed_entries = set()
        self.setup_ui()
        self.update_idletasks()
        self.focus_force()
        self.after(100, self.focus_force)
        self.bind("<Configure>", self.on_window_resize)

    def setup_ui(self):
        # Responsive PanedWindow for main/processed/errors
        self.paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(self.paned)
        self.main_frame = ttk.Frame(self.notebook)
        self.processed_frame = ttk.Frame(self.notebook)
        self.error_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text="Main")
        self.notebook.add(self.processed_frame, text="Processed")
        self.notebook.add(self.error_frame, text="Errors")
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.paned.add(self.notebook, weight=3)

        # Error tab
        self.error_listbox = tk.Listbox(self.error_frame)
        self.error_listbox.pack(fill=tk.BOTH, expand=True)

        # Bottom bar for progress and bulk action buttons
        bottom_bar = ttk.Frame(self)
        bottom_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(bottom_bar, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=5, pady=5)

        # Status and bulk action buttons
        button_frame = ttk.Frame(bottom_bar)
        button_frame.pack(side=tk.RIGHT, padx=5)

        self.status_label = ttk.Label(button_frame, text="Ready")
        self.status_label.pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame, text="Update selected", command=lambda: self.bulk_action("update")
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            button_frame, text="Create selected", command=lambda: self.bulk_action("create")
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            button_frame, text="Remove selected", command=lambda: self.bulk_action("remove")
        ).pack(side=tk.LEFT, padx=2)

        # Add import .ics button at the top
        import_bar = ttk.Frame(self)
        import_bar.pack(side=tk.TOP, fill=tk.X)
        ttk.Button(import_bar, text="Import .ics file", command=self.import_ics_file).pack(
            side=tk.LEFT, padx=5, pady=5
        )

        # Scrollable frame for entries (Main)
        canvas = tk.Canvas(self.main_frame)
        scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=canvas.yview)
        self.entries_frame = ttk.Frame(canvas)
        self.entries_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.entries_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Use an "active canvas" approach to capture mousewheel/trackpad reliably on macOS
        self._active_scroll_canvas = None

        def on_enter(e, cvs=canvas):
            # set active canvas and bind platform mouse events
            self._active_scroll_canvas = cvs
            # bind global events so trackpad gestures are captured
            self.bind_all("<MouseWheel>", self._on_mousewheel)
            self.bind_all("<Button-4>", self._on_mousewheel)
            self.bind_all("<Button-5>", self._on_mousewheel)

        def on_leave(e):
            self._active_scroll_canvas = None
            # unbind global events
            try:
                self.unbind_all("<MouseWheel>")
                self.unbind_all("<Button-4>")
                self.unbind_all("<Button-5>")
            except Exception:
                pass

        canvas.bind("<Enter>", on_enter)
        canvas.bind("<Leave>", on_leave)

        # Scrollable frame for processed entries
        proc_canvas = tk.Canvas(self.processed_frame)
        proc_scrollbar = ttk.Scrollbar(
            self.processed_frame, orient="vertical", command=proc_canvas.yview
        )
        self.processed_entries_frame = ttk.Frame(proc_canvas)
        self.processed_entries_frame.bind(
            "<Configure>", lambda e: proc_canvas.configure(scrollregion=proc_canvas.bbox("all"))
        )
        proc_canvas.create_window((0, 0), window=self.processed_entries_frame, anchor="nw")
        proc_canvas.configure(yscrollcommand=proc_scrollbar.set)
        proc_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        proc_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.populate_entries()
        self.populate_processed_entries()

    def _on_mousewheel(self, event):
        # Scroll the currently active canvas (set on Enter)
        canvas = getattr(self, "_active_scroll_canvas", None)
        if not canvas:
            return
        # platform-specific behavior
        sys = platform.system()
        if sys == "Windows":
            # Windows: event.delta is multiple of 120
            try:
                steps = int(-1 * (event.delta / 120))
            except Exception:
                steps = -1 if event.delta > 0 else 1
            canvas.yview_scroll(steps, "units")
        elif sys == "Darwin":
            # macOS: event.delta is small/continuous; use delta directly
            # scale down so scrolling isn't too fast
            try:
                # delta on mac is typically small floats; invert sign for natural direction
                steps = int(-1 * event.delta)
            except Exception:
                steps = -1 if getattr(event, "delta", 0) > 0 else 1
            # If steps is 0 (very small delta), fall back to +/-1
            if steps == 0:
                steps = -1 if getattr(event, "delta", 0) > 0 else 1
            canvas.yview_scroll(steps, "units")
        else:
            # Linux: Button-4 / Button-5
            if getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")

    def import_ics_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("ICS files", "*.ics")])
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                c = Calendar(f.read())
            new_rows = []
            for event in c.events:
                title = event.name
                start = event.begin.format("YYYY-MM-DD")
                # Check for duplicates in CSV
                if not ((self.df["Title"] == title) & (self.df["Start"] == start)).any():
                    new_rows.append(
                        {"Title": title, "Start": start, "done": False, "removed": False}
                    )
            if new_rows:
                self.df = pd.concat([self.df, pd.DataFrame(new_rows)], ignore_index=True)
                # Ensure all rows have 'done' and 'removed' columns
                for col in ["done", "removed"]:
                    if col not in self.df.columns:
                        self.df[col] = False
                    self.df[col] = self.df[col].fillna(False).astype(bool)
                self.df.to_csv(CSV_PATH, index=False)
                messagebox.showinfo("Import", f"Imported {len(new_rows)} new events.")
                self.populate_entries()
                self.populate_processed_entries()
                self.update_idletasks()
            else:
                messagebox.showinfo("Import", "No new events found in .ics file.")
        except Exception as e:
            messagebox.showerror("Import Error", str(e))

    def fuzzy_match(self, calendar_title):
        matches = process.extract(calendar_title, self.contact_names, scorer=fuzz.WRatio, limit=10)
        matches = sorted(matches, key=lambda x: x[1], reverse=True)
        return [m[0] for m in matches if m[1] > 60] or ["<Create new contact>"]

    def populate_entries(self):
        # Clear previous
        for widget in self.entries_frame.winfo_children():
            widget.destroy()
        self.entries = []
        # Detect repeated titles (same title, different dates)
        title_counts = self.df.groupby("Title")["Start"].nunique()
        repeated_titles = set(title_counts[title_counts > 1].index)

        for idx, row in self.df.iterrows():
            if row.get("done", False) or row.get("removed", False):
                continue
            entry = {}
            entry["idx"] = idx
            entry["title"] = row["Title"]
            entry["date"] = row["Start"]
            entry["selected"] = tk.BooleanVar(value=False)
            entry["match_list"] = self.fuzzy_match(row["Title"])
            entry["match_var"] = tk.StringVar(value=entry["match_list"][0])

            # Main frame
            entry["frame"] = frame = ttk.Frame(self.entries_frame)
            frame.pack(fill=tk.X, padx=2, pady=2)

            # Configure grid weights
            frame.grid_columnconfigure(1, weight=1)  # Title column expands

            # Highlight repeated titles
            if row["Title"] in repeated_titles:
                frame.config(style="Repeated.TFrame")

            # Left side: checkbox and title
            ttk.Checkbutton(frame, variable=entry["selected"]).grid(
                row=0, column=0, padx=2, sticky="w"
            )
            ttk.Label(frame, text=f"{row['Title']} ({row['Start']})").grid(
                row=0, column=1, padx=5, sticky="w"
            )

            # Center: match dropdown
            entry["combobox"] = cb = ttk.Combobox(
                frame,
                values=entry["match_list"],
                textvariable=entry["match_var"],
                width=20,
                state="readonly",
            )
            cb.grid(row=0, column=2, padx=2, sticky="w")

            # Search section
            search_frame = ttk.Frame(frame)
            search_frame.grid(row=0, column=3, padx=5, sticky="w")

            ttk.Label(search_frame, text="Search:").grid(row=0, column=0, padx=2)
            entry["search_var"] = tk.StringVar()

            def set_match(val, e=entry):
                e["match_var"].set(val)

            autocomplete = AutocompleteEntry(
                self.contact_names, entry["search_var"], search_frame, set_match, width=15
            )
            autocomplete.grid(row=0, column=1, padx=2)
            entry["autocomplete"] = autocomplete

            # Right side: action buttons
            entry["right_frame"] = right_frame = ttk.Frame(frame)
            right_frame.grid(row=0, column=4, padx=5, sticky="e")

            # Individual action buttons
            update_btn = ttk.Button(
                right_frame, text="Update", command=lambda e=entry: self.set_and_update_entry(e)
            )
            update_btn.grid(row=0, column=0, padx=1)
            entry["update_btn"] = update_btn

            new_btn = ttk.Button(
                right_frame, text="New", command=lambda e=entry: self.create_new_contact_ui(e)
            )
            new_btn.grid(row=0, column=1, padx=1)
            entry["new_btn"] = new_btn

            remove_btn = ttk.Button(
                right_frame, text="Remove", command=lambda e=entry: self.remove_entry(e)
            )
            remove_btn.grid(row=0, column=2, padx=1)
            entry["remove_btn"] = remove_btn

            self.entries.append(entry)

            # Apply responsive layout
            self.update_entry_layout(entry)

        # Style for repeated events
        style = ttk.Style()
        style.configure("Repeated.TFrame", background="#ffe4e1")

    def on_window_resize(self, event=None):
        if event and event.widget == self:
            # Update all entry rows for responsiveness
            for entry in self.entries:
                self.update_entry_layout(entry)

    def update_entry_layout(self, entry):
        # Check if we need to show compact layout
        window_width = self.winfo_width()
        is_compact = window_width < 1000

        if is_compact and not hasattr(entry, "_is_compact"):
            # Switch to compact layout
            self.make_entry_compact(entry)
        elif not is_compact and hasattr(entry, "_is_compact"):
            # Switch to full layout
            self.make_entry_full(entry)

    def make_entry_compact(self, entry):
        # Hide individual buttons, show more menu
        if "update_btn" in entry:
            entry["update_btn"].grid_remove()
            entry["new_btn"].grid_remove()
            entry["remove_btn"].grid_remove()

        if "more_btn" not in entry:
            # Create more menu
            more_menu = tk.Menu(entry["right_frame"], tearoff=0)
            more_menu.add_command(
                label="Update", command=lambda e=entry: self.set_and_update_entry(e)
            )
            more_menu.add_command(
                label="New", command=lambda e=entry: self.create_new_contact_ui(e)
            )
            more_menu.add_command(label="Remove", command=lambda e=entry: self.remove_entry(e))

            more_btn = ttk.Menubutton(entry["right_frame"], text="⋯", width=3)
            more_btn["menu"] = more_menu
            more_btn.grid(row=0, column=3, padx=2)
            entry["more_btn"] = more_btn
        else:
            entry["more_btn"].grid()

        entry["_is_compact"] = True

    def make_entry_full(self, entry):
        # Show individual buttons, hide more menu
        if "more_btn" in entry:
            entry["more_btn"].grid_remove()

        if "update_btn" in entry:
            entry["update_btn"].grid()
            entry["new_btn"].grid()
            entry["remove_btn"].grid()

        if hasattr(entry, "_is_compact"):
            delattr(entry, "_is_compact")

    def bulk_action(self, action):
        selected_entries = [e for e in self.entries if e["selected"].get()]
        if not selected_entries:
            messagebox.showinfo("Info", "No entries selected.")
            return

        if action == "remove":
            # Bulk remove - mark as removed and refresh UI
            for entry in selected_entries:
                idx = entry["idx"]
                self.df.loc[idx, "removed"] = True
                entry["frame"].destroy()
            self.df.to_csv(CSV_PATH, index=False)
            return

        # For update and create actions, use threading
        self.progress = 0
        self.total = len(selected_entries)
        self.status_label.config(text=f"Processing {action}...")
        self.progress_var.set(0)

        if action == "update":
            threading.Thread(
                target=self._update_contacts_thread, args=(selected_entries,), daemon=True
            ).start()
        elif action == "create":
            threading.Thread(
                target=self._create_contacts_thread, args=(selected_entries,), daemon=True
            ).start()

    def _create_contacts_thread(self, selected_entries):
        for i, entry in enumerate(selected_entries):
            idx = entry["idx"]
            try:
                create_contact(self.service, entry["title"], entry["date"])
                self.df.loc[idx, "done"] = True
                self.df.to_csv(CSV_PATH, index=False)
                self.after(0, self.move_entry_to_processed, entry)
            except Exception as e:
                self.errors.append(f"{entry['title']}: {str(e)}")
                self.error_listbox.insert(tk.END, f"{entry['title']}: {str(e)}")
            self.progress = i + 1
            self.progress_var.set(100 * self.progress / self.total)
            self.status_label.config(text=f"Created {self.progress}/{self.total}")
        self.status_label.config(text="Done")

    # Remove the old update_selected_contacts method and rename _update_contacts_thread
    def _update_contacts_thread(self, selected_entries):
        for i, entry in enumerate(selected_entries):
            contact_name = entry["match_var"].get()
            idx = entry["idx"]
            try:
                if contact_name == "<Create new contact>":
                    create_contact(self.service, entry["title"], entry["date"])
                else:
                    contact_obj = next(
                        (
                            c
                            for c in self.contacts
                            if c.get("names") and c["names"][0]["displayName"] == contact_name
                        ),
                        None,
                    )
                    if contact_obj:
                        full_contact = get_contact_details(
                            self.service, contact_obj["resourceName"]
                        )
                        update_birthday(self.service, full_contact, entry["date"])
                    else:
                        raise Exception("Contact not found")
                self.df.loc[idx, "done"] = True
                self.df.to_csv(CSV_PATH, index=False)
                self.after(0, self.move_entry_to_processed, entry)
            except Exception as e:
                self.errors.append(f"{entry['title']}: {str(e)}")
                self.error_listbox.insert(tk.END, f"{entry['title']}: {str(e)}")
            self.progress = i + 1
            self.progress_var.set(100 * self.progress / self.total)
            self.status_label.config(text=f"Updated {self.progress}/{self.total}")
        self.status_label.config(text="Done")

    def move_entry_to_processed(self, entry):
        # Remove from main entries frame
        entry["frame"].destroy()
        self.populate_processed_entries()

    def remove_entry(self, entry):
        idx = entry["idx"]
        self.df.loc[idx, "removed"] = True
        self.df.to_csv(CSV_PATH, index=False)
        entry["frame"].destroy()

    def populate_processed_entries(self):
        for widget in self.processed_entries_frame.winfo_children():
            widget.destroy()
        self.processed_entries = []
        for idx, row in self.df.iterrows():
            if not row.get("done", False):
                continue
            frame = ttk.Frame(self.processed_entries_frame)
            frame.pack(fill=tk.X, padx=2, pady=2)
            ttk.Label(frame, text=f"{row['Title']} ({row['Start']})").pack(side=tk.LEFT, padx=5)
            ttk.Label(frame, text="Processed").pack(side=tk.LEFT, padx=5)
            self.processed_entries.append(frame)

    def set_and_update_entry(self, entry):
        # Immediately update contact for this entry and mark as processed
        contact_name = entry["match_var"].get()
        idx = entry["idx"]
        try:
            if contact_name == "<Create new contact>":
                create_contact(self.service, entry["title"], entry["date"])
            else:
                contact_obj = next(
                    (
                        c
                        for c in self.contacts
                        if c.get("names") and c["names"][0]["displayName"] == contact_name
                    ),
                    None,
                )
                if contact_obj:
                    full_contact = get_contact_details(self.service, contact_obj["resourceName"])
                    update_birthday(self.service, full_contact, entry["date"])
                else:
                    raise Exception("Contact not found")
            self.df.loc[idx, "done"] = True
            self.df.to_csv(CSV_PATH, index=False)
            self.move_entry_to_processed(entry)
        except Exception as e:
            self.errors.append(f"{entry['title']}: {str(e)}")
            self.error_listbox.insert(tk.END, f"{entry['title']}: {str(e)}")

    def create_new_contact_ui(self, entry):
        try:
            create_contact(self.service, entry["title"], entry["date"])
            self.df.loc[entry["idx"], "done"] = True
            self.df.to_csv(CSV_PATH, index=False)
            self.move_entry_to_processed(entry)
        except Exception as e:
            self.errors.append(f"{entry['title']}: {str(e)}")
            self.error_listbox.insert(tk.END, f"{entry['title']}: {str(e)}")


def main():
    # Try to load CSV, create empty DataFrame if missing or empty
    if not os.path.exists(CSV_PATH) or os.path.getsize(CSV_PATH) == 0:
        df = pd.DataFrame(columns=["Title", "Start", "done", "removed"])
        df.to_csv(CSV_PATH, index=False)
    else:
        df = pd.read_csv(CSV_PATH)
        # Ensure required columns exist and fill missing values
        for col in ["Title", "Start"]:
            if col not in df.columns:
                df[col] = ""
        for col in ["done", "removed"]:
            if col not in df.columns:
                df[col] = False
            df[col] = df[col].fillna(False).astype(bool)
        df.to_csv(CSV_PATH, index=False)
    creds = authenticate_google()
    service = build("people", "v1", credentials=creds)
    contacts = get_contacts(service)
    app = CalendarSyncApp(df, contacts, service)
    app.mainloop()


if __name__ == "__main__":
    main()
