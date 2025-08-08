
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd

APP_TITLE = "Email Stage Labeler"
WINDOW_SIZE = "1200x800"

# Column indices (0-based): L=11, M=12
COL_EMAIL = 11
COL_STAGE = 12

# Key bindings to stages
STAGES = {
    "Reviewing": ("<Up>", "w", "W"),
    "Passed": ("<Down>", "s", "S"),
    "Bounceback": ("<Left>", "a", "A"),
    "Auto-Reply": ("<Right>", "d", "D"),
}

class EmailStageLabeler(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(WINDOW_SIZE)
        self.minsize(900, 600)

        self.df = None
        self.csv_path = None
        # row_indices holds only rows where stage is blank
        self.row_indices = []
        # pointer into row_indices
        self.i = 0

        self._build_ui()
        self._bind_keys()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=0)  # sidebar
        self.grid_columnconfigure(1, weight=1)  # content
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

        # Sidebar
        self.sidebar = tk.Frame(self, padx=12, pady=12)
        self.sidebar.grid(row=0, column=0, sticky="nsw")
        self._build_sidebar()

        # Content
        self.content = tk.Frame(self, padx=12, pady=12)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(1, weight=1)

        # Header
        self.header = tk.Frame(self.content)
        self.header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.header.grid_columnconfigure(0, weight=1)

        self.file_label = tk.Label(self.header, text="No file loaded", anchor="w", font=("Segoe UI", 10, "italic"))
        self.file_label.grid(row=0, column=0, sticky="w")

        self.progress_label = tk.Label(self.header, text="", anchor="e", font=("Segoe UI", 10, "bold"))
        self.progress_label.grid(row=0, column=1, sticky="e", padx=(8, 0))

        self.stage_label = tk.Label(self.header, text="", anchor="e", fg="#444", font=("Segoe UI", 10))
        self.stage_label.grid(row=0, column=2, sticky="e", padx=(8, 0))

        # Email display
        self.email_text = ScrolledText(self.content, wrap=tk.WORD, font=("Consolas", 14), undo=False)
        self.email_text.grid(row=1, column=0, sticky="nsew")
        self.email_text.configure(state="disabled")

        # Controls
        self.controls = tk.Frame(self)
        self.controls.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=12)
        self.controls.grid_columnconfigure(1, weight=1)

        self.open_btn = tk.Button(self.controls, text="Open CSV", command=self.open_csv, width=16)
        self.open_btn.grid(row=0, column=0, padx=(0, 8))

        self.back_btn = tk.Button(self.controls, text="Back", command=self.go_back, width=12, state="disabled")
        self.back_btn.grid(row=0, column=2, padx=4)

        self.skip_btn = tk.Button(self.controls, text="Skip", command=self.skip_row, width=12, state="disabled")
        self.skip_btn.grid(row=0, column=3, padx=4)

        self.quit_btn = tk.Button(self.controls, text="Quit", command=self.on_quit, width=12)
        self.quit_btn.grid(row=0, column=4, padx=4)

        # Status bar
        self.status = tk.StringVar(value="Open a CSV to begin")
        self.status_bar = tk.Label(self, textvariable=self.status, anchor="w", relief=tk.SUNKEN, bd=1)
        self.status_bar.grid(row=2, column=0, columnspan=2, sticky="ew")

    def _build_sidebar(self):
        title = tk.Label(self.sidebar, text="Key Bindings", font=("Segoe UI", 12, "bold"))
        title.pack(anchor="w", pady=(0, 6))

        bindings_text = [
            ("Up or W", "Reviewing"),
            ("Down or S", "Passed"),
            ("Left or A", "Bounceback"),
            ("Right or D", "Auto-Reply"),
            ("Back button", "Go to previous row"),
            ("Skip button", "Skip without labeling"),
        ]

        for keys, action in bindings_text:
            row = tk.Frame(self.sidebar)
            row.pack(anchor="w", pady=2, fill="x")
            tk.Label(row, text=f"{keys}:", font=("Segoe UI", 10, "bold")).pack(side="left")
            tk.Label(row, text=f" {action}", font=("Segoe UI", 10)).pack(side="left")

    def _bind_keys(self):
        # Stage hotkeys
        self.bind("<Up>", lambda e: self.set_stage("Reviewing"))
        self.bind("<w>", lambda e: self.set_stage("Reviewing"))
        self.bind("<W>", lambda e: self.set_stage("Reviewing"))

        self.bind("<Down>", lambda e: self.set_stage("Passed"))
        self.bind("<s>", lambda e: self.set_stage("Passed"))
        self.bind("<S>", lambda e: self.set_stage("Passed"))

        self.bind("<Left>", lambda e: self.set_stage("Bounceback"))
        self.bind("<a>", lambda e: self.set_stage("Bounceback"))
        self.bind("<A>", lambda e: self.set_stage("Bounceback"))

        self.bind("<Right>", lambda e: self.set_stage("Auto-Reply"))
        self.bind("<d>", lambda e: self.set_stage("Auto-Reply"))
        self.bind("<D>", lambda e: self.set_stage("Auto-Reply"))

    def open_csv(self):
        path = filedialog.askopenfilename(
            title="Select CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            df = pd.read_csv(path, header=0, dtype=str, encoding="utf-8-sig", engine="python")
        except Exception as e:
            try:
                df = pd.read_csv(path, header=0, dtype=str, encoding="utf-8", engine="python", errors="replace")
            except Exception as e2:
                messagebox.showerror("Error", f"Failed to read CSV:\n{e}\n{e2}")
                return

        # Ensure required columns exist
        max_needed = max(COL_EMAIL, COL_STAGE)
        if df.shape[1] <= max_needed:
            messagebox.showerror("Error", f"CSV must have at least {max_needed+1} columns so that L and M exist")
            return

        # Normalize stage column to string dtype
        df.iloc[:, COL_STAGE] = df.iloc[:, COL_STAGE].astype("string")

        self.df = df
        self.csv_path = path

        # Build list of only unlabeled rows (stage blank or NaN)
        self.row_indices = [
            idx for idx in range(len(self.df))
            if pd.isna(self.df.iat[idx, COL_STAGE]) or str(self.df.iat[idx, COL_STAGE]).strip() == ""
        ]
        self.i = 0

        self.file_label.config(text=os.path.basename(self.csv_path))
        self.back_btn.config(state="normal")
        self.skip_btn.config(state="normal")

        if not self.row_indices:
            self.status.set("All rows already labeled")
            self._set_email_text("All rows already labeled")
            self.progress_label.config(text="0 / 0")
            self.stage_label.config(text="")
        else:
            self.status.set(f"Loaded CSV. {len(self.row_indices)} unlabeled rows found")
            self.show_current()

    def show_current(self):
        if self.df is None:
            return
        if not self.row_indices:
            self._set_email_text("All rows already labeled")
            self.progress_label.config(text="0 / 0")
            self.stage_label.config(text="")
            return

        # Clamp pointer
        self.i = max(0, min(self.i, len(self.row_indices) - 1))
        idx = self.row_indices[self.i]

        email_val = self._safe_val(self.df.iat[idx, COL_EMAIL])
        stage_val = self._safe_val(self.df.iat[idx, COL_STAGE])

        self._set_email_text(email_val if email_val is not None else "")
        self.progress_label.config(text=f"{self.i + 1} / {len(self.row_indices)}")
        self.stage_label.config(text=f"Current Stage: {stage_val if stage_val else '(blank)'}")

    def _set_email_text(self, text):
        self.email_text.configure(state="normal")
        self.email_text.delete("1.0", tk.END)
        self.email_text.insert(tk.END, text)
        self.email_text.configure(state="disabled")

    def _safe_val(self, v):
        if pd.isna(v):
            return ""
        return str(v)

    def set_stage(self, stage_name):
        if self.df is None or not self.row_indices:
            return

        # Dataframe index of current row
        idx = self.row_indices[self.i]

        # Set stage and save
        self.df.iat[idx, COL_STAGE] = stage_name
        self._save_csv()

        # Remove this index from the unlabeled list, since it's now labeled
        self.row_indices.pop(self.i)

        if not self.row_indices:
            # No more items
            self.status.set("Saved. All rows labeled")
            self.show_current()
            return

        # Keep pointer at same position to show next item (which has shifted into this index)
        if self.i >= len(self.row_indices):
            self.i = len(self.row_indices) - 1

        self.status.set("Saved")
        self.show_current()

    def skip_row(self):
        if self.df is None or not self.row_indices:
            return
        if self.i < len(self.row_indices) - 1:
            self.i += 1
        self.show_current()

    def go_back(self):
        if self.df is None or not self.row_indices:
            return
        if self.i > 0:
            self.i -= 1
        self.show_current()

    def _save_csv(self):
        if self.df is None or not self.csv_path:
            return
        try:
            # Direct overwrite
            self.df.to_csv(self.csv_path, index=False, encoding="utf-8-sig")
            self.status.set("Saved")
        except Exception as e:
            self.status.set(f"Save failed: {e}")

    def on_quit(self):
        self.destroy()


def main():
    app = EmailStageLabeler()
    app.mainloop()


if __name__ == "__main__":
    main()
