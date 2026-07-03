"""
RO Code Report Merger — Desktop App
------------------------------------
Upload an Excel sheet, and any RO Code with multiple rows will
automatically be merged into a single row.

To run:
    pip install -r requirements.txt
    python desktop_app.py

To build a Windows .exe:
    pip install pyinstaller
    pyinstaller --noconsole --onefile --name "RO_Code_Merger" desktop_app.py
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

GROUP_KEY = "RO Code"

# Ye columns SAME rehte hain har RO Code ke liye (pehla non-empty value liya jayega)
SAME_COLUMNS = [
    "Zone", "Region", "Salesarea", "Vendor", "RO Code",
    "RO Sap Code", "RO Name", "ROC Date",
]

# Ye columns DIFFERENT hote hain, unique values comma (,) se jode jayenge
DIFFERENT_COLUMNS = [
    "Ticket No", "Ageing Days", "Device", "Device Details",
    "Current Dependency", "Automation Vendor Comments", "HPCL Comments",
]

APP_BG = "#F4F6F9"
PRIMARY = "#1B4F72"
ACCENT = "#2E86C1"
SUCCESS = "#1E8449"
CARD_BG = "#FFFFFF"

FONT_TITLE = ("Segoe UI", 22, "bold")
FONT_SUBTITLE = ("Segoe UI", 12)
FONT_BODY = ("Segoe UI", 11)
FONT_BUTTON = ("Segoe UI", 12, "bold")


def merge_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Merges duplicate RO Code rows into a single row."""
    if GROUP_KEY not in df.columns:
        raise ValueError(f"'{GROUP_KEY}' column not found in the sheet.")

    same_cols = [c for c in SAME_COLUMNS if c in df.columns and c != GROUP_KEY]
    diff_cols = [c for c in DIFFERENT_COLUMNS if c in df.columns]

    def merge_group(g: pd.DataFrame) -> pd.Series:
        result = {GROUP_KEY: g.name}
        for col in same_cols:
            vals = g[col].dropna()
            result[col] = vals.iloc[0] if not vals.empty else None
        for col in diff_cols:
            vals = g[col].dropna().astype(str).unique().tolist()
            result[col] = ", ".join(vals)
        result["Ticket Count"] = len(g)
        return pd.Series(result)

    merged = (
        df.groupby(GROUP_KEY, dropna=False, sort=False)
        .apply(merge_group)
        .reset_index(drop=True)
    )
    final_cols = [GROUP_KEY] + same_cols + diff_cols + ["Ticket Count"]
    return merged[final_cols]


class ROMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RO Code Report Merger")
        self.geometry("1150x700")
        self.minsize(950, 600)
        self.configure(fg_color=APP_BG)

        self.source_df = None
        self.merged_df = None
        self.file_path = None

        self._build_header()
        self._build_upload_card()
        self._build_status_bar()
        self._build_table_area()

    # ------------------------------------------------------------------ UI
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=PRIMARY, height=90, corner_radius=0)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        title = ctk.CTkLabel(
            header, text="📋  RO Code Report Merger",
            font=FONT_TITLE, text_color="white",
        )
        title.pack(anchor="w", padx=30, pady=(14, 0))

        subtitle = ctk.CTkLabel(
            header,
            text="Merge duplicate RO Code rows into a single clean row — ready report for engineers",
            font=FONT_SUBTITLE, text_color="#D6EAF8",
        )
        subtitle.pack(anchor="w", padx=30, pady=(0, 12))

    def _build_upload_card(self):
        card = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=12)
        card.pack(fill="x", padx=25, pady=18)

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=18)

        self.upload_btn = ctk.CTkButton(
            inner, text="📂  Upload Excel File", font=FONT_BUTTON,
            fg_color=ACCENT, hover_color=PRIMARY, height=42, width=200,
            command=self.upload_file,
        )
        self.upload_btn.grid(row=0, column=0, padx=(0, 15))

        self.file_label = ctk.CTkLabel(
            inner, text="No file selected",
            font=FONT_BODY, text_color="#555",
        )
        self.file_label.grid(row=0, column=1, sticky="w")

        self.merge_btn = ctk.CTkButton(
            inner, text="🔀  Merge", font=FONT_BUTTON,
            fg_color=SUCCESS, hover_color="#145A32", height=42, width=160,
            state="disabled", command=self.run_merge_thread,
        )
        self.merge_btn.grid(row=0, column=2, padx=(15, 0))

        self.export_btn = ctk.CTkButton(
            inner, text="⬇️  Export Report", font=FONT_BUTTON,
            fg_color=PRIMARY, hover_color="#0B2E4E", height=42, width=170,
            state="disabled", command=self.export_file,
        )
        self.export_btn.grid(row=0, column=3, padx=(15, 0))

        inner.grid_columnconfigure(1, weight=1)

    def _build_status_bar(self):
        self.status_label = ctk.CTkLabel(
            self, text="Ready. Upload a file above to get started.",
            font=FONT_BODY, text_color="#444", anchor="w",
        )
        self.status_label.pack(fill="x", padx=30, pady=(0, 8))

        self.progress = ctk.CTkProgressBar(self, mode="indeterminate")
        self.progress.set(0)

    def _build_table_area(self):
        table_card = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=12)
        table_card.pack(fill="both", expand=True, padx=25, pady=(0, 20))

        self.table_title = ctk.CTkLabel(
            table_card, text="Preview", font=("Segoe UI", 14, "bold"),
            text_color=PRIMARY, anchor="w",
        )
        self.table_title.pack(anchor="w", padx=18, pady=(14, 6))

        tree_frame = tk.Frame(table_card, bg=CARD_BG)
        tree_frame.pack(fill="both", expand=True, padx=18, pady=(0, 16))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Custom.Treeview", background="white", fieldbackground="white",
            foreground="#222", rowheight=26, font=("Segoe UI", 10),
        )
        style.configure(
            "Custom.Treeview.Heading", font=("Segoe UI", 10, "bold"),
            background=ACCENT, foreground="white",
        )
        style.map("Custom.Treeview", background=[("selected", "#AED6F1")])

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        self.tree = ttk.Treeview(
            tree_frame, style="Custom.Treeview",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    # ------------------------------------------------------------- ACTIONS
    def upload_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if not path:
            return
        try:
            self.status_label.configure(text="Loading file...")
            self.update_idletasks()
            self.source_df = pd.read_excel(path, sheet_name=0)
            self.file_path = path
            self.file_label.configure(text=os.path.basename(path))
            self.merge_btn.configure(state="normal")
            self.export_btn.configure(state="disabled")
            self.status_label.configure(
                text=f"File loaded ✅  ({len(self.source_df)} rows). "
                     f"Click 'Merge' to continue."
            )
            self._populate_table(self.source_df.head(50), "Original Data (preview - top 50 rows)")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file:\n{e}")
            self.status_label.configure(text="File load failed.")

    def run_merge_thread(self):
        self.merge_btn.configure(state="disabled")
        self.upload_btn.configure(state="disabled")
        self.status_label.configure(text="Merging in progress...")
        self.progress.pack(fill="x", padx=30, pady=(0, 10))
        self.progress.start()
        threading.Thread(target=self._do_merge, daemon=True).start()

    def _do_merge(self):
        try:
            merged = merge_dataframe(self.source_df)
            self.merged_df = merged
            self.after(0, self._on_merge_done, None)
        except Exception as e:
            self.after(0, self._on_merge_done, e)

    def _on_merge_done(self, error):
        self.progress.stop()
        self.progress.pack_forget()
        self.upload_btn.configure(state="normal")
        self.merge_btn.configure(state="normal")

        if error:
            messagebox.showerror("Error", f"Something went wrong during merge:\n{error}")
            self.status_label.configure(text="Merge failed.")
            return

        self.export_btn.configure(state="normal")
        self.status_label.configure(
            text=f"Merge complete ✅  {len(self.source_df)} rows → "
                 f"{len(self.merged_df)} unique RO Codes."
        )
        self._populate_table(self.merged_df, "Merged Report (preview)")

    def export_file(self):
        if self.merged_df is None:
            return
        default_name = "merged_report.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save merged report",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        try:
            self.merged_df.to_excel(path, index=False, sheet_name="Merged Report")
            self.status_label.configure(text=f"Report saved ✅  → {path}")
            messagebox.showinfo("Saved", f"Merged report saved successfully:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file:\n{e}")

    def _populate_table(self, df: pd.DataFrame, title: str):
        self.table_title.configure(text=title)
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in df.columns:
            self.tree.heading(col, text=str(col))
            width = min(max(len(str(col)) * 9, 90), 220)
            self.tree.column(col, width=width, anchor="w")

        for _, row in df.iterrows():
            values = ["" if pd.isna(v) else str(v) for v in row]
            self.tree.insert("", "end", values=values)


if __name__ == "__main__":
    app = ROMergerApp()
    app.mainloop()