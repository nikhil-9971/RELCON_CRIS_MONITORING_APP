"""
RO Code Report Merger — Desktop App
------------------------------------
Upload an Excel sheet, and any RO Code with multiple rows will
automatically be merged into a single row.

To run:
    pip install -r requirements.txt
    python desktop_app.py

To build a Windows .exe (with the app icon bundled):
    pip install pyinstaller
    pyinstaller --noconsole --onefile --icon=icon.ico ^
        --add-data "icon.ico;." --name "RO_Code_Merger" desktop_app.py
"""

import os
import sys
import ctypes
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd


# ---------------------------------------------------------------------------
# Icon / taskbar identity helpers
# ---------------------------------------------------------------------------
def resource_path(relative_path: str) -> str:
    """
    Resolves a bundled resource's path, both when running as a normal
    Python script and when running as a PyInstaller --onefile .exe
    (where files are unpacked into a temporary _MEIPASS folder).
    """
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def set_windows_app_id(app_id: str = "Relcon.ROCodeMerger.1"):
    """
    Without this, Windows groups the app under the generic Python icon
    when it is pinned to the taskbar. Setting a unique AppUserModelID
    tells Windows this is its own distinct application.
    """
    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# CONFIG — business logic
# ---------------------------------------------------------------------------
GROUP_KEY = "RO Code"

# Columns that stay the SAME for a given RO Code (first non-empty value is kept)
SAME_COLUMNS = [
    "Zone", "Region", "Salesarea", "Vendor", "RO Code",
    "RO Sap Code", "RO Name", "ROC Date",
]

# Columns that DIFFER per ticket — unique values are joined with a comma
DIFFERENT_COLUMNS = [
    "Ticket No", "Ageing Days", "Device", "Device Details",
    "Current Dependency", "Automation Vendor Comments", "HPCL Comments",
]

# ---------------------------------------------------------------------------
# CONFIG — visual design system
# ---------------------------------------------------------------------------
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

COLOR = {
    "bg":            "#EEF1F6",
    "surface":       "#FFFFFF",
    "border":        "#E2E6ED",
    "sidebar":       "#0F1B33",
    "sidebar_soft":  "#1B2C4F",
    "primary":       "#2151D6",
    "primary_hover": "#1A3FAE",
    "success":       "#12875A",
    "success_hover": "#0C6B45",
    "danger":        "#D64545",
    "text":          "#101828",
    "text_muted":    "#5B6472",
    "text_faint":    "#8A94A6",
    "row_alt":       "#F6F8FB",
    "chip_blue_bg":  "#E8EEFD",
    "chip_blue_fg":  "#2151D6",
    "chip_green_bg": "#E5F6EE",
    "chip_green_fg": "#12875A",
    "chip_amber_bg": "#FDF3E3",
    "chip_amber_fg": "#B4740E",
}

F_DISPLAY  = ("Segoe UI Semibold", 20)
F_H2       = ("Segoe UI Semibold", 14)
F_BODY     = ("Segoe UI", 12)
F_BODY_SM  = ("Segoe UI", 11)
F_LABEL    = ("Segoe UI Semibold", 11)
F_BUTTON   = ("Segoe UI Semibold", 12)
F_MONO_NUM = ("Segoe UI Semibold", 22)
F_SIDEBAR_TITLE = ("Segoe UI Semibold", 15)
F_SIDEBAR_STEP  = ("Segoe UI Semibold", 12)
F_SIDEBAR_DESC  = ("Segoe UI", 10)


# ---------------------------------------------------------------------------
# Business logic
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# Reusable UI components
# ---------------------------------------------------------------------------
class StatCard(ctk.CTkFrame):
    """Small metric card, e.g. 'Total Rows: 198'."""

    def __init__(self, master, label, chip_bg, chip_fg, **kwargs):
        super().__init__(
            master, fg_color=COLOR["surface"], corner_radius=10,
            border_width=1, border_color=COLOR["border"], **kwargs,
        )
        self.value_label = ctk.CTkLabel(
            self, text="—", font=F_MONO_NUM, text_color=chip_fg,
        )
        self.value_label.pack(anchor="w", padx=18, pady=(16, 0))

        ctk.CTkLabel(
            self, text=label, font=F_LABEL, text_color=COLOR["text_muted"],
        ).pack(anchor="w", padx=18, pady=(2, 16))

        self._chip_bg = chip_bg

    def set_value(self, value):
        self.value_label.configure(text=str(value))


class StepRow(ctk.CTkFrame):
    """One entry in the sidebar step tracker."""

    def __init__(self, master, number, title, desc, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        self.circle = ctk.CTkLabel(
            self, text=str(number), width=28, height=28, corner_radius=14,
            fg_color=COLOR["sidebar_soft"], text_color="#8FA3D6",
            font=F_SIDEBAR_STEP,
        )
        self.circle.grid(row=0, column=0, rowspan=2, sticky="n", padx=(0, 12))

        self.title_label = ctk.CTkLabel(
            self, text=title, font=F_SIDEBAR_STEP, text_color="#8FA3D6",
            anchor="w", justify="left",
        )
        self.title_label.grid(row=0, column=1, sticky="w")

        self.desc_label = ctk.CTkLabel(
            self, text=desc, font=F_SIDEBAR_DESC, text_color="#57699C",
            anchor="w", justify="left", wraplength=170,
        )
        self.desc_label.grid(row=1, column=1, sticky="w", pady=(2, 0))

        self.grid_columnconfigure(1, weight=1)

    def set_state(self, state):
        """state: 'pending' | 'active' | 'done'"""
        if state == "active":
            self.circle.configure(fg_color=COLOR["primary"], text_color="white")
            self.title_label.configure(text_color="white")
            self.desc_label.configure(text_color="#B7C4EA")
        elif state == "done":
            self.circle.configure(fg_color=COLOR["success"], text_color="white")
            self.circle.configure(text="✓")
            self.title_label.configure(text_color="#D6ECDF")
            self.desc_label.configure(text_color="#7FAF95")
        else:
            self.circle.configure(fg_color=COLOR["sidebar_soft"], text_color="#8FA3D6")
            self.title_label.configure(text_color="#8FA3D6")
            self.desc_label.configure(text_color="#57699C")


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------
class ROMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RO Code Report Merger")
        self.geometry("1280x760")
        self.minsize(1080, 640)
        self.configure(fg_color=COLOR["bg"])
        self._set_app_icon()

        self.source_df = None
        self.merged_df = None
        self.file_path = None

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main_area()
        self._set_step("upload")

    def _set_app_icon(self):
        """Sets the window/taskbar icon (icon.ico) if it's available."""
        try:
            ico_path = resource_path("icon.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
                # Some Windows taskbar/pin scenarios also read iconphoto;
                # setting both keeps the icon consistent everywhere.
                icon_img = tk.PhotoImage(file=resource_path("icon.png")) \
                    if os.path.exists(resource_path("icon.png")) else None
                if icon_img is not None:
                    self.iconphoto(True, icon_img)
                    self._icon_ref = icon_img  # prevent garbage collection
        except Exception:
            pass

    # ------------------------------------------------------------ SIDEBAR
    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, fg_color=COLOR["sidebar"], corner_radius=0, width=250)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)

        brand = ctk.CTkFrame(sidebar, fg_color="transparent")
        brand.pack(fill="x", padx=24, pady=(28, 30))

        ctk.CTkLabel(brand, text="⛽", font=("Segoe UI", 26)).pack(anchor="w")
        ctk.CTkLabel(
            brand, text="RO Code Merger", font=F_SIDEBAR_TITLE, text_color="white",
        ).pack(anchor="w", pady=(6, 0))
        ctk.CTkLabel(
            brand, text="Ticket Consolidation Tool", font=F_SIDEBAR_DESC,
            text_color="#7C8CBB",
        ).pack(anchor="w", pady=(2, 0))

        ctk.CTkFrame(sidebar, fg_color="#233257", height=1).pack(fill="x", padx=24, pady=(0, 26))

        steps_wrap = ctk.CTkFrame(sidebar, fg_color="transparent")
        steps_wrap.pack(fill="x", padx=24)

        self.step_upload = StepRow(
            steps_wrap, 1, "Upload File",
            "Select the raw Excel export from your system",
        )
        self.step_upload.pack(fill="x", pady=(0, 22))

        self.step_merge = StepRow(
            steps_wrap, 2, "Merge Records",
            "Combine duplicate RO Code rows into one",
        )
        self.step_merge.pack(fill="x", pady=(0, 22))

        self.step_export = StepRow(
            steps_wrap, 3, "Export Report",
            "Save the clean report as an Excel file",
        )
        self.step_export.pack(fill="x", pady=(0, 22))

        footer = ctk.CTkFrame(sidebar, fg_color="transparent")
        footer.pack(side="bottom", fill="x", padx=24, pady=20)
        ctk.CTkFrame(sidebar, fg_color="#233257", height=1).pack(
            side="bottom", fill="x", padx=24, pady=(0, 4)
        )
        ctk.CTkLabel(
            footer, text="v1.0.0", font=F_SIDEBAR_DESC, text_color="#57699C",
        ).pack(anchor="w")

    def _set_step(self, step: str):
        mapping = {
            "upload": ("active", "pending", "pending"),
            "merged": ("done", "active", "pending"),
            "exported": ("done", "done", "done"),
        }
        u, m, e = mapping[step]
        self.step_upload.set_state(u)
        self.step_merge.set_state(m)
        self.step_export.set_state(e)

    # --------------------------------------------------------- MAIN AREA
    def _build_main_area(self):
        main = ctk.CTkFrame(self, fg_color=COLOR["bg"], corner_radius=0)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)

        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=32, pady=(28, 4))
        ctk.CTkLabel(
            header, text="Report Merger", font=F_DISPLAY, text_color=COLOR["text"],
        ).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="Merge duplicate RO Code tickets into a single, engineer-ready row.",
            font=F_BODY, text_color=COLOR["text_muted"],
        ).pack(anchor="w", pady=(2, 0))

        card = ctk.CTkFrame(
            main, fg_color=COLOR["surface"], corner_radius=14,
            border_width=1, border_color=COLOR["border"],
        )
        card.grid(row=1, column=0, sticky="ew", padx=32, pady=18)
        card.grid_columnconfigure(0, weight=1)

        self.dropzone = ctk.CTkButton(
            card, text="📂   Click to select an Excel file  (.xlsx / .xls)",
            font=F_BUTTON, fg_color=COLOR["chip_blue_bg"], text_color=COLOR["chip_blue_fg"],
            hover_color="#DCE6FB", corner_radius=10, height=64,
            border_width=2, border_color=COLOR["primary"], border_spacing=0,
            command=self.upload_file, anchor="w",
        )
        self.dropzone.grid(row=0, column=0, columnspan=3, sticky="ew", padx=20, pady=(20, 8))

        self.file_label = ctk.CTkLabel(
            card, text="No file selected yet", font=F_BODY_SM, text_color=COLOR["text_faint"],
            anchor="w",
        )
        self.file_label.grid(row=1, column=0, columnspan=3, sticky="w", padx=22, pady=(0, 16))

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.grid(row=2, column=0, columnspan=3, sticky="ew", padx=20, pady=(0, 20))

        self.merge_btn = ctk.CTkButton(
            btn_row, text="🔀   Merge Duplicates", font=F_BUTTON,
            fg_color=COLOR["success"], hover_color=COLOR["success_hover"],
            height=42, width=190, corner_radius=8,
            state="disabled", command=self.run_merge_thread,
        )
        self.merge_btn.pack(side="left")

        self.export_btn = ctk.CTkButton(
            btn_row, text="⬇   Export to Excel", font=F_BUTTON,
            fg_color=COLOR["primary"], hover_color=COLOR["primary_hover"],
            height=42, width=180, corner_radius=8,
            state="disabled", command=self.export_file,
        )
        self.export_btn.pack(side="left", padx=(12, 0))

        self.reset_btn = ctk.CTkButton(
            btn_row, text="Reset", font=F_BUTTON,
            fg_color="transparent", text_color=COLOR["text_muted"],
            hover_color=COLOR["row_alt"], border_width=1, border_color=COLOR["border"],
            height=42, width=90, corner_radius=8,
            command=self.reset_app,
        )
        self.reset_btn.pack(side="right")

        self.status_label = ctk.CTkLabel(
            btn_row, text="Ready", font=F_BODY_SM, text_color=COLOR["text_faint"],
        )
        self.status_label.pack(side="right", padx=(0, 16))

        self.progress = ctk.CTkProgressBar(card, mode="indeterminate", height=4)

        self.stats_row = ctk.CTkFrame(main, fg_color="transparent")
        self.stats_row.grid(row=2, column=0, sticky="ew", padx=32, pady=(0, 6))
        self.stats_row.grid_columnconfigure((0, 1, 2), weight=1, uniform="stats")

        self.stat_total = StatCard(self.stats_row, "Total Ticket Rows", *self._chip("blue"))
        self.stat_unique = StatCard(self.stats_row, "Unique RO Codes", *self._chip("green"))
        self.stat_merged = StatCard(self.stats_row, "Rows Consolidated", *self._chip("amber"))
        self.stat_total.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.stat_unique.grid(row=0, column=1, sticky="ew", padx=8)
        self.stat_merged.grid(row=0, column=2, sticky="ew", padx=(8, 0))
        self.stats_row.grid_remove()

        table_card = ctk.CTkFrame(
            main, fg_color=COLOR["surface"], corner_radius=14,
            border_width=1, border_color=COLOR["border"],
        )
        table_card.grid(row=3, column=0, sticky="nsew", padx=32, pady=(6, 28))
        table_card.grid_columnconfigure(0, weight=1)
        table_card.grid_rowconfigure(1, weight=1)

        title_row = ctk.CTkFrame(table_card, fg_color="transparent")
        title_row.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 8))
        self.table_title = ctk.CTkLabel(
            title_row, text="Preview", font=F_H2, text_color=COLOR["text"],
        )
        self.table_title.pack(side="left")
        self.table_subtitle = ctk.CTkLabel(
            title_row, text="Upload a file to see data here", font=F_BODY_SM,
            text_color=COLOR["text_faint"],
        )
        self.table_subtitle.pack(side="left", padx=(10, 0))

        self._build_table(table_card)

    @staticmethod
    def _chip(kind):
        return {
            "blue": (COLOR["chip_blue_bg"], COLOR["chip_blue_fg"]),
            "green": (COLOR["chip_green_bg"], COLOR["chip_green_fg"]),
            "amber": (COLOR["chip_amber_bg"], COLOR["chip_amber_fg"]),
        }[kind]

    def _build_table(self, parent):
        tree_frame = tk.Frame(parent, bg=COLOR["surface"])
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Pro.Treeview",
            background=COLOR["surface"], fieldbackground=COLOR["surface"],
            foreground=COLOR["text"], rowheight=30, font=F_BODY_SM,
            borderwidth=0,
        )
        style.configure(
            "Pro.Treeview.Heading",
            font=F_LABEL, background="#F8FAFC", foreground=COLOR["text_muted"],
            borderwidth=0, relief="flat",
        )
        style.map("Pro.Treeview.Heading", background=[("active", "#F0F3F9")])
        style.map(
            "Pro.Treeview",
            background=[("selected", "#DCE6FB")],
            foreground=[("selected", COLOR["text"])],
        )
        style.layout("Pro.Treeview", [("Pro.Treeview.treearea", {"sticky": "nswe"})])

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        self.tree = ttk.Treeview(
            tree_frame, style="Pro.Treeview",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree.tag_configure("odd", background=COLOR["surface"])
        self.tree.tag_configure("even", background=COLOR["row_alt"])

    # ------------------------------------------------------------- ACTIONS
    def upload_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if not path:
            return
        try:
            self.status_label.configure(text="Loading file…")
            self.update_idletasks()
            self.source_df = pd.read_excel(path, sheet_name=0)
            self.file_path = path
            self.file_label.configure(
                text=f"📄  {os.path.basename(path)}   ·   {len(self.source_df)} rows loaded",
                text_color=COLOR["text_muted"],
            )
            self.dropzone.configure(text="📂   Change file")
            self.merge_btn.configure(state="normal")
            self.export_btn.configure(state="disabled")
            self.stats_row.grid_remove()
            self.status_label.configure(text="File loaded ✓", text_color=COLOR["success"])
            self._set_step("upload")
            self.table_subtitle.configure(text="Original data — top 50 rows")
            self._populate_table(self.source_df.head(50))
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file:\n{e}")
            self.status_label.configure(text="File load failed", text_color=COLOR["danger"])

    def run_merge_thread(self):
        self.merge_btn.configure(state="disabled")
        self.dropzone.configure(state="disabled")
        self.status_label.configure(text="Merging…", text_color=COLOR["text_faint"])
        self.progress.grid(row=3, column=0, columnspan=3, sticky="ew", padx=20, pady=(0, 4))
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
        self.progress.grid_forget()
        self.dropzone.configure(state="normal")
        self.merge_btn.configure(state="normal")

        if error:
            messagebox.showerror("Error", f"Something went wrong during merge:\n{error}")
            self.status_label.configure(text="Merge failed", text_color=COLOR["danger"])
            return

        self.export_btn.configure(state="normal")
        self.status_label.configure(text="Merge complete ✓", text_color=COLOR["success"])
        self._set_step("merged")

        total = len(self.source_df)
        unique = len(self.merged_df)
        self.stat_total.set_value(total)
        self.stat_unique.set_value(unique)
        self.stat_merged.set_value(total - unique)
        self.stats_row.grid()

        self.table_subtitle.configure(text=f"Merged report — {unique} rows")
        self._populate_table(self.merged_df)

    def export_file(self):
        if self.merged_df is None:
            return
        path = filedialog.asksaveasfilename(
            title="Save merged report",
            defaultextension=".xlsx",
            initialfile="merged_report.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        try:
            self.merged_df.to_excel(path, index=False, sheet_name="Merged Report")
            self.status_label.configure(text="Report saved ✓", text_color=COLOR["success"])
            self._set_step("exported")
            messagebox.showinfo("Saved", f"Merged report saved successfully:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file:\n{e}")

    def reset_app(self):
        self.source_df = None
        self.merged_df = None
        self.file_path = None
        self.file_label.configure(text="No file selected yet", text_color=COLOR["text_faint"])
        self.dropzone.configure(text="📂   Click to select an Excel file  (.xlsx / .xls)")
        self.merge_btn.configure(state="disabled")
        self.export_btn.configure(state="disabled")
        self.status_label.configure(text="Ready", text_color=COLOR["text_faint"])
        self.stats_row.grid_remove()
        self.table_title.configure(text="Preview")
        self.table_subtitle.configure(text="Upload a file to see data here")
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []
        self._set_step("upload")

    def _populate_table(self, df: pd.DataFrame):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in df.columns:
            self.tree.heading(col, text=str(col))
            width = min(max(len(str(col)) * 9, 90), 220)
            self.tree.column(col, width=width, anchor="w")

        for i, (_, row) in enumerate(df.iterrows()):
            values = ["" if pd.isna(v) else str(v) for v in row]
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=values, tags=(tag,))


if __name__ == "__main__":
    set_windows_app_id()
    app = ROMergerApp()
    app.mainloop()
