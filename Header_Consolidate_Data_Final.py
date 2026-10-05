import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re
import difflib
from collections import OrderedDict


# ============================================================
# HEADER CONSOLIDATE DATA
# ============================================================
#
# Features
# 1. Select an Excel workbook.
# 2. Select/preview any worksheet in a popup.
# 3. Excel-like filtering in the preview popup.
# 4. Optional filtered rows can be used for consolidation.
# 5. Read all worksheet headers.
# 6. Automatically suggest equivalent headers with confidence.
# 7. Manually map headers using colors.
# 8. Same color + different header names = warning only.
# 9. No person/record matching, deduplication, or row merging.
# 10. Every source row is preserved unless the user explicitly
#     chooses to consolidate only filtered rows.
# 11. Add "Source Sheet" to every consolidated row.
#
# Requirements:
#     pip install pandas openpyxl
#
# ============================================================


APP_TITLE = "Header Consolidate Data"

COLORS = OrderedDict([
    ("None", "#FFFFFF"),
    ("Yellow", "#FFF2CC"),
    ("Blue", "#9DC3E6"),
    ("Green", "#A9D18E"),
    ("Orange", "#F4B183"),
    ("Pink", "#F4CCCC"),
    ("Purple", "#D9D2E9"),
    ("Red", "#EA9999"),
    ("Light Blue", "#CFE2F3"),
    ("Light Green", "#D9EAD3"),
    ("Grey", "#D9D9D9"),
])

# Colors used only for displaying automatic suggestions.
AUTO_GROUP_COLORS = [
    "Yellow", "Blue", "Green", "Orange", "Pink",
    "Purple", "Red", "Light Blue", "Light Green", "Grey"
]

# Common header aliases. This is intentionally conservative.
HEADER_ALIASES = {
    "documentid": "Document ID",
    "documentno": "Document ID",
    "documentnumber": "Document ID",
    "docid": "Document ID",
    "docno": "Document ID",
    "docnumber": "Document ID",

    "name": "Name",
    "fullname": "Name",
    "individualname": "Name",
    "personname": "Name",
    "customername": "Name",
    "membername": "Name",

    "firstname": "First Name",
    "fname": "First Name",

    "middlename": "Middle Name",
    "mname": "Middle Name",

    "lastname": "Last Name",
    "lname": "Last Name",
    "surname": "Last Name",
    "familyname": "Last Name",

    "tin": "TIN",
    "tinnumber": "TIN",
    "taxnumber": "TIN",
    "taxid": "TIN",
    "taxidentificationnumber": "TIN",

    "ssn": "SSN",
    "ssnnumber": "SSN",
    "socialsecuritynumber": "SSN",

    "dob": "DOB",
    "dateofbirth": "DOB",
    "birthdate": "DOB",
    "birthdt": "DOB",

    "providerid": "Provider ID",
    "providernumber": "Provider ID",
    "providerno": "Provider ID",
    "providernum": "Provider ID",

    "dl": "DL",
    "dlnumber": "DL",
    "driverslicense": "DL",
    "driverslicensenumber": "DL",
    "drivinglicensenumber": "DL",

    "phone": "Phone",
    "phonenumber": "Phone",
    "mobilenumber": "Phone",
    "mobile": "Phone",
    "telephone": "Phone",

    "email": "Email",
    "emailaddress": "Email",

    "address": "Address",
    "address1": "Address",
    "streetaddress": "Address",

    "city": "City",
    "state": "State",
    "zipcode": "ZIP Code",
    "zip": "ZIP Code",
    "postalcode": "ZIP Code",
    "postcode": "ZIP Code",

    "gender": "Gender",
    "sex": "Gender",

    "status": "Status",
}


def normalize_header(value):
    """Normalize a header for automatic comparison."""
    text = "" if value is None else str(value)
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9]+", "", text)
    return text


def clean_display_value(value):
    if pd.isna(value):
        return ""
    return str(value)


def header_similarity(a, b):
    """Return a 0-100 similarity score."""
    na = normalize_header(a)
    nb = normalize_header(b)

    if not na or not nb:
        return 0.0

    if na == nb:
        return 100.0

    if na in HEADER_ALIASES and HEADER_ALIASES[na] == HEADER_ALIASES.get(nb, ""):
        return 100.0

    if HEADER_ALIASES.get(na) and HEADER_ALIASES.get(na) == b:
        return 98.0

    if HEADER_ALIASES.get(nb) and HEADER_ALIASES.get(nb) == a:
        return 98.0

    # Compare canonical alias names.
    ca = HEADER_ALIASES.get(na, na)
    cb = HEADER_ALIASES.get(nb, nb)

    if ca == cb:
        return 96.0

    ratio = difflib.SequenceMatcher(None, na, nb).ratio() * 100
    canonical_ratio = difflib.SequenceMatcher(
        None,
        normalize_header(ca),
        normalize_header(cb)
    ).ratio() * 100

    # A modest boost for meaningful token overlap.
    token_a = set(re.findall(r"[a-z]+", str(a).lower()))
    token_b = set(re.findall(r"[a-z]+", str(b).lower()))

    overlap = 0.0
    if token_a and token_b:
        overlap = (
            len(token_a & token_b) /
            max(len(token_a | token_b), 1)
        ) * 100

    return max(ratio, canonical_ratio, overlap)


def best_header_match(header, all_unique_headers):
    """Return (suggestion, confidence)."""
    if not all_unique_headers:
        return header, 100.0

    best = None
    best_score = -1

    for candidate in all_unique_headers:
        score = header_similarity(header, candidate)
        if score > best_score:
            best_score = score
            best = candidate

    # Avoid presenting weak fuzzy guesses as strong matches.
    if best_score < 55:
        return "", best_score

    return best, best_score


class FilterPopup(tk.Toplevel):
    """
    Excel-style-ish worksheet filter popup.

    The popup provides:
      - worksheet selector
      - global search
      - per-column filter dropdown
      - select all / clear
      - filtered row count
      - checkbox to use filtered rows for consolidation
    """

    def __init__(self, parent, sheet_data, initial_sheet=None,
                 filtered_rows_by_sheet=None):
        super().__init__(parent)

        self.parent = parent
        self.sheet_data = sheet_data
        self.filtered_rows_by_sheet = filtered_rows_by_sheet or {}

        self.title("Select Sheet & Filter Data")
        self.geometry("1400x800")
        self.minsize(1050, 650)
        self.transient(parent)
        self.grab_set()

        self.current_sheet = initial_sheet or (
            next(iter(sheet_data)) if sheet_data else None
        )

        self.current_df = pd.DataFrame()
        self.base_df = pd.DataFrame()
        self.filtered_df = pd.DataFrame()

        self.filter_values = {}
        self.column_filters = {}
        self.column_filter_vars = {}

        self.search_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self.use_filter_var = tk.BooleanVar(value=False)

        self.result = None

        self.create_ui()

        if self.current_sheet:
            self.load_sheet(self.current_sheet)

        self.protocol("WM_DELETE_WINDOW", self.cancel)

    def create_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Label(
            top,
            text="Sheet:",
            font=("Segoe UI", 10, "bold")
        ).pack(side="left")

        self.sheet_var = tk.StringVar(value=self.current_sheet or "")

        self.sheet_combo = ttk.Combobox(
            top,
            textvariable=self.sheet_var,
            values=list(self.sheet_data.keys()),
            state="readonly",
            width=35
        )
        self.sheet_combo.pack(side="left", padx=(7, 15))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_changed)

        ttk.Label(
            top,
            text="Search all columns:"
        ).pack(side="left")

        search_entry = ttk.Entry(
            top,
            textvariable=self.search_var,
            width=35
        )
        search_entry.pack(side="left", padx=7)
        search_entry.bind("<KeyRelease>", lambda e: self.apply_filters())

        ttk.Button(
            top,
            text="Clear All Filters",
            command=self.clear_all_filters
        ).pack(side="left", padx=5)

        ttk.Button(
            top,
            text="Refresh",
            command=self.apply_filters
        ).pack(side="left", padx=5)

        # Filter controls.
        filter_frame = ttk.LabelFrame(
            self,
            text="Column Filters",
            padding=8
        )
        filter_frame.pack(fill="x", padx=10, pady=(0, 8))

        self.filter_canvas = tk.Canvas(
            filter_frame,
            height=85,
            highlightthickness=0
        )
        self.filter_scroll = ttk.Scrollbar(
            filter_frame,
            orient="horizontal",
            command=self.filter_canvas.xview
        )

        self.filter_inner = ttk.Frame(self.filter_canvas)
        self.filter_window = self.filter_canvas.create_window(
            (0, 0),
            window=self.filter_inner,
            anchor="nw"
        )

        self.filter_inner.bind(
            "<Configure>",
            lambda e: self.filter_canvas.configure(
                scrollregion=self.filter_canvas.bbox("all")
            )
        )

        self.filter_canvas.configure(
            xscrollcommand=self.filter_scroll.set
        )

        self.filter_canvas.pack(
            fill="x",
            expand=True,
            side="top"
        )
        self.filter_scroll.pack(
            fill="x",
            side="bottom"
        )

        # Data table.
        table_frame = ttk.Frame(self, padding=(10, 0, 10, 0))
        table_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            table_frame,
            show="headings",
            selectmode="browse"
        )

        sy = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )
        sx = ttk.Scrollbar(
            table_frame,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=sy.set,
            xscrollcommand=sx.set
        )

        self.tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Bottom controls.
        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")

        ttk.Label(
            bottom,
            textvariable=self.status_var
        ).pack(side="left")

        ttk.Checkbutton(
            bottom,
            text="Use filtered rows for FINAL CONSOLIDATE",
            variable=self.use_filter_var
        ).pack(side="left", padx=25)

        ttk.Button(
            bottom,
            text="Cancel",
            command=self.cancel
        ).pack(side="right", padx=5)

        ttk.Button(
            bottom,
            text="Apply / Continue",
            command=self.apply_and_close
        ).pack(side="right", padx=5)

    def on_sheet_changed(self, event=None):
        sheet = self.sheet_var.get()
        if sheet:
            self.load_sheet(sheet)

    def load_sheet(self, sheet):
        self.current_sheet = sheet
        self.base_df = self.sheet_data[sheet].copy()

        if sheet in self.filtered_rows_by_sheet:
            previous = self.filtered_rows_by_sheet[sheet]
            if len(previous) == len(self.base_df):
                self.use_filter_var.set(False)

        self.search_var.set("")
        self.column_filters = {}
        self.build_filter_controls()
        self.apply_filters()

    def build_filter_controls(self):
        for widget in self.filter_inner.winfo_children():
            widget.destroy()

        self.column_filter_vars = {}

        for index, column in enumerate(self.base_df.columns):
            cell = ttk.Frame(
                self.filter_inner,
                relief="groove",
                padding=5
            )
            cell.grid(row=0, column=index, padx=4, pady=2, sticky="n")

            ttk.Label(
                cell,
                text=str(column),
                font=("Segoe UI", 9, "bold")
            ).pack(anchor="w")

            values = self.base_df[column].map(
                clean_display_value
            ).drop_duplicates().tolist()

            # Keep a practical list while still allowing all values.
            values = sorted(
                values,
                key=lambda x: x.lower()
            )

            combo_values = ["(All)"] + values

            var = tk.StringVar(value="(All)")
            combo = ttk.Combobox(
                cell,
                textvariable=var,
                values=combo_values,
                state="readonly",
                width=22
            )
            combo.pack()
            combo.bind(
                "<<ComboboxSelected>>",
                lambda e: self.apply_filters()
            )

            self.column_filter_vars[str(column)] = var

    def apply_filters(self):
        if self.base_df is None or self.base_df.empty:
            self.filtered_df = self.base_df.copy()
            self.populate_table()
            self.status_var.set("0 rows")
            return

        df = self.base_df.copy()

        # Global search.
        search_text = self.search_var.get().strip().lower()
        if search_text:
            mask = pd.Series(False, index=df.index)

            for column in df.columns:
                mask = mask | df[column].map(
                    clean_display_value
                ).str.lower().str.contains(
                    re.escape(search_text),
                    na=False
                )

            df = df[mask]

        # Per-column filters.
        for column, var in self.column_filter_vars.items():
            selected = var.get()

            if selected and selected != "(All)":
                df = df[
                    df[column].map(clean_display_value) == selected
                ]

        self.filtered_df = df
        self.populate_table()

        self.status_var.set(
            f"Sheet: {self.current_sheet}    "
            f"Showing {len(df):,} of {len(self.base_df):,} rows"
        )

    def populate_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        columns = [str(c) for c in self.base_df.columns]
        self.tree["columns"] = columns

        for column in columns:
            self.tree.heading(column, text=column)
            self.tree.column(
                column,
                width=150,
                minwidth=90,
                anchor="w"
            )

        for _, row in self.filtered_df.iterrows():
            values = [
                clean_display_value(row[c])
                for c in self.base_df.columns
            ]
            self.tree.insert("", "end", values=values)

    def clear_all_filters(self):
        self.search_var.set("")

        for var in self.column_filter_vars.values():
            var.set("(All)")

        self.apply_filters()

    def apply_and_close(self):
        if self.current_sheet:
            self.filtered_rows_by_sheet[self.current_sheet] = (
                self.filtered_df.copy()
            )

        self.result = {
            "selected_sheet": self.current_sheet,
            "filtered_rows_by_sheet": self.filtered_rows_by_sheet,
            "use_filtered_rows": self.use_filter_var.get()
        }

        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


class HeaderMappingPopup(tk.Toplevel):
    """
    Displays all headers and automatic suggestions before the
    main color-mapping interface is used.
    """

    def __init__(self, parent, sheet_headers):
        super().__init__(parent)

        self.parent = parent
        self.sheet_headers = sheet_headers

        self.title("Automatic Header Identification")
        self.geometry("1250x700")
        self.minsize(950, 550)
        self.transient(parent)
        self.grab_set()

        self.create_ui()

    def create_ui(self):
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")

        ttk.Label(
            top,
            text="Automatic Header Identification",
            font=("Segoe UI", 16, "bold")
        ).pack(anchor="w")

        ttk.Label(
            top,
            text=(
                "The program compares header names across all sheets. "
                "Review the suggested equivalent header and confidence "
                "before using the manual color mapping."
            ),
            wraplength=1100
        ).pack(anchor="w", pady=(5, 10))

        frame = ttk.Frame(self, padding=(12, 0, 12, 0))
        frame.pack(fill="both", expand=True)

        columns = (
            "sheet",
            "column",
            "original",
            "suggestion",
            "confidence"
        )

        tree = ttk.Treeview(
            frame,
            columns=columns,
            show="headings"
        )

        headings = {
            "sheet": "Sheet",
            "column": "Column",
            "original": "Original Header",
            "suggestion": "Automatically Identified As",
            "confidence": "Confidence"
        }

        widths = {
            "sheet": 190,
            "column": 90,
            "original": 280,
            "suggestion": 300,
            "confidence": 130
        }

        for col in columns:
            tree.heading(col, text=headings[col])
            tree.column(col, width=widths[col], anchor="w")

        sy = ttk.Scrollbar(
            frame,
            orient="vertical",
            command=tree.yview
        )
        sx = ttk.Scrollbar(
            frame,
            orient="horizontal",
            command=tree.xview
        )

        tree.configure(
            yscrollcommand=sy.set,
            xscrollcommand=sx.set
        )

        tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        unique_headers = []
        for headers in self.sheet_headers.values():
            for info in headers:
                h = info["original_header"]
                if h not in unique_headers:
                    unique_headers.append(h)

        for sheet_name, headers in self.sheet_headers.items():
            for info in headers:
                suggestion, score = best_header_match(
                    info["original_header"],
                    unique_headers
                )

                # Don't suggest itself as a different mapping.
                if suggestion == info["original_header"]:
                    display_suggestion = suggestion
                elif not suggestion:
                    display_suggestion = "No reliable match"
                else:
                    display_suggestion = suggestion

                tree.insert(
                    "",
                    "end",
                    values=(
                        sheet_name,
                        self.parent.column_letter(
                            info["column_index"]
                        ),
                        info["original_header"],
                        display_suggestion,
                        f"{score:.1f}%"
                    )
                )

        bottom = ttk.Frame(self, padding=12)
        bottom.pack(fill="x")

        ttk.Label(
            bottom,
            text=(
                "Automatic identification is a suggestion only. "
                "You can still assign the final colors manually."
            )
        ).pack(side="left")

        ttk.Button(
            bottom,
            text="Close",
            command=self.destroy
        ).pack(side="right")


class HeaderConsolidatorApp:
    # ========================================================
    # INITIALIZATION
    # ========================================================

    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1250x800")
        self.root.minsize(1050, 700)

        self.file_path = None
        self.output_path = None

        self.sheet_headers = OrderedDict()
        self.tree_metadata = {}
        self.selected_item = None

        # Full data for every worksheet.
        self.sheet_data = OrderedDict()

        # Optional filtered data selected in filter popup.
        self.filtered_rows_by_sheet = OrderedDict()

        # If True, FINAL CONSOLIDATE uses filtered rows.
        self.use_filtered_rows = False

        self.create_styles()
        self.create_ui()

    # ========================================================
    # STYLES
    # ========================================================

    def create_styles(self):
        style = ttk.Style()

        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 20, "bold")
        )

        style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 10)
        )

        style.configure(
            "Header.TLabel",
            font=("Segoe UI", 11, "bold")
        )

        style.configure(
            "Action.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=8
        )

    # ========================================================
    # CREATE UI
    # ========================================================

    def create_ui(self):
        top_frame = ttk.Frame(self.root, padding=15)
        top_frame.pack(fill="x")

        ttk.Label(
            top_frame,
            text="Header Consolidate Data",
            style="Title.TLabel"
        ).pack(anchor="w")

        ttk.Label(
            top_frame,
            text=(
                "Automatically review headers, filter worksheet data, "
                "map corresponding headers using colors, and consolidate "
                "all rows with a Source Sheet reference."
            ),
            style="Subtitle.TLabel"
        ).pack(anchor="w", pady=(3, 10))

        # File selection.
        file_frame = ttk.LabelFrame(
            self.root,
            text="1. Select Excel File",
            padding=12
        )
        file_frame.pack(
            fill="x",
            padx=15,
            pady=(0, 8)
        )

        self.file_label = ttk.Label(
            file_frame,
            text="No Excel file selected.",
            foreground="gray"
        )
        self.file_label.pack(
            side="left",
            fill="x",
            expand=True
        )

        ttk.Button(
            file_frame,
            text="Select Excel File",
            command=self.select_file,
            style="Action.TButton"
        ).pack(side="right")

        # Sheet/filter controls.
        data_frame = ttk.LabelFrame(
            self.root,
            text="2. Sheet Selection & Excel-Style Data Filter",
            padding=10
        )
        data_frame.pack(
            fill="x",
            padx=15,
            pady=(0, 8)
        )

        self.selected_sheet_var = tk.StringVar(
            value="No sheet selected"
        )

        ttk.Label(
            data_frame,
            textvariable=self.selected_sheet_var,
            font=("Segoe UI", 10, "bold")
        ).pack(side="left", padx=(0, 15))

        ttk.Button(
            data_frame,
            text="Open Sheet & Filter Popup",
            command=self.open_filter_popup
        ).pack(side="left", padx=5)

        ttk.Button(
            data_frame,
            text="View Automatic Header Identification",
            command=self.open_auto_header_popup
        ).pack(side="left", padx=5)

        self.filter_status_var = tk.StringVar(
            value="All rows will be used unless filtering is enabled."
        )

        ttk.Label(
            data_frame,
            textvariable=self.filter_status_var,
            foreground="gray"
        ).pack(side="right")

        # Mapping area.
        instruction_frame = ttk.LabelFrame(
            self.root,
            text="3. Header Color Mapping",
            padding=10
        )
        instruction_frame.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=(0, 8)
        )

        instruction = (
            "Assign the SAME color to headers that represent the same "
            "logical field across sheets. Automatic header identification "
            "is shown separately as a suggestion.\n\n"
            "Example: Document ID and Document Number can both be Yellow.\n\n"
            "If the same color is assigned to different header names, "
            "the program shows a warning. You can still continue."
        )

        ttk.Label(
            instruction_frame,
            text=instruction,
            justify="left"
        ).pack(anchor="w", pady=(0, 8))

        table_frame = ttk.Frame(instruction_frame)
        table_frame.pack(fill="both", expand=True)

        columns = (
            "sheet",
            "column",
            "header",
            "auto",
            "confidence",
            "color"
        )

        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse"
        )

        headings = {
            "sheet": "Sheet",
            "column": "Column",
            "header": "Header Name",
            "auto": "Auto Identified As",
            "confidence": "Confidence",
            "color": "Assigned Color"
        }

        for col in columns:
            self.tree.heading(col, text=headings[col])

        self.tree.column("sheet", width=170)
        self.tree.column("column", width=75)
        self.tree.column("header", width=250)
        self.tree.column("auto", width=250)
        self.tree.column("confidence", width=100)
        self.tree.column("color", width=140)

        scrollbar_y = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )

        scrollbar_x = ttk.Scrollbar(
            table_frame,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )

        self.tree.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        scrollbar_y.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        scrollbar_x.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree.bind(
            "<<TreeviewSelect>>",
            self.on_header_select
        )

        # Color controls.
        color_frame = ttk.Frame(instruction_frame)
        color_frame.pack(fill="x", pady=(10, 0))

        ttk.Label(
            color_frame,
            text="Selected Header:"
        ).pack(side="left")

        self.selected_header_label = ttk.Label(
            color_frame,
            text="None",
            font=("Segoe UI", 10, "bold")
        )
        self.selected_header_label.pack(
            side="left",
            padx=(5, 20)
        )

        ttk.Label(
            color_frame,
            text="Color:"
        ).pack(side="left")

        self.color_var = tk.StringVar(value="None")

        self.color_combo = ttk.Combobox(
            color_frame,
            textvariable=self.color_var,
            values=list(COLORS.keys()),
            state="readonly",
            width=18
        )
        self.color_combo.pack(
            side="left",
            padx=8
        )

        self.color_combo.bind(
            "<<ComboboxSelected>>",
            self.apply_color
        )

        ttk.Button(
            color_frame,
            text="Apply Color",
            command=self.apply_color
        ).pack(side="left")

        # Bottom buttons.
        bottom_frame = ttk.Frame(
            self.root,
            padding=(15, 0, 15, 15)
        )
        bottom_frame.pack(fill="x")

        ttk.Button(
            bottom_frame,
            text="Validate Mapping",
            command=self.validate_mapping,
            style="Action.TButton"
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            bottom_frame,
            text="Clear Mapping",
            command=self.clear_mapping
        ).pack(side="left")

        ttk.Button(
            bottom_frame,
            text="FINAL CONSOLIDATE",
            command=self.final_consolidate,
            style="Action.TButton"
        ).pack(side="right")

        self.status_var = tk.StringVar(
            value="Please select an Excel file."
        )

        status_label = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief="sunken",
            anchor="w",
            padding=5
        )

        status_label.pack(
            side="bottom",
            fill="x"
        )

    # ========================================================
    # SELECT EXCEL FILE
    # ========================================================

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel Workbook",
            filetypes=[
                ("Excel Files", "*.xlsx *.xlsm"),
                ("Excel Workbook", "*.xlsx"),
                ("Excel Macro Workbook", "*.xlsm"),
                ("All Files", "*.*")
            ]
        )

        if not path:
            return

        self.file_path = path

        self.file_label.config(
            text=os.path.basename(path),
            foreground="black"
        )

        self.load_workbook_headers()

    # ========================================================
    # LOAD WORKBOOK
    # ========================================================

    def load_workbook_headers(self):
        try:
            self.status_var.set("Reading workbook...")
            self.root.update_idletasks()

            excel = pd.ExcelFile(self.file_path)

            self.sheet_headers.clear()
            self.sheet_data.clear()
            self.filtered_rows_by_sheet.clear()

            for sheet_name in excel.sheet_names:
                df = pd.read_excel(
                    self.file_path,
                    sheet_name=sheet_name,
                    dtype=object
                )

                self.sheet_data[sheet_name] = df

                headers = []
                used_names = {}

                for column_index, value in enumerate(df.columns):
                    header = str(value).strip()

                    if not header:
                        header = (
                            f"Unnamed Column {column_index + 1}"
                        )

                    if header in used_names:
                        used_names[header] += 1

                        display_header = (
                            f"{header} "
                            f"(Duplicate {used_names[header]})"
                        )
                    else:
                        used_names[header] = 1
                        display_header = header

                    headers.append({
                        "column_index": column_index,
                        "header": display_header,
                        "original_header": header,
                        "color": "None"
                    })

                self.sheet_headers[sheet_name] = headers

            self.populate_tree()

            total_headers = sum(
                len(headers)
                for headers in self.sheet_headers.values()
            )

            self.status_var.set(
                f"Loaded {len(self.sheet_headers)} sheet(s) "
                f"and {total_headers} header(s)."
            )

            if self.sheet_headers:
                first_sheet = next(
                    iter(self.sheet_headers.keys())
                )
                self.selected_sheet_var.set(
                    f"Selected sheet: {first_sheet}"
                )

                self.filter_status_var.set(
                    "All rows will be used. "
                    "Use the Sheet & Filter popup if required."
                )

                # Show automatic identification immediately.
                self.root.after(
                    150,
                    self.open_auto_header_popup
                )

        except Exception as e:
            messagebox.showerror(
                "Error Reading Excel",
                "Unable to read the workbook.\n\n"
                f"{type(e).__name__}: {e}"
            )

            self.status_var.set(
                "Error reading workbook."
            )

    # ========================================================
    # AUTO HEADER SUGGESTIONS
    # ========================================================

    def get_auto_suggestion(self, header):
        all_unique_headers = []

        for headers in self.sheet_headers.values():
            for info in headers:
                h = info["original_header"]

                if h not in all_unique_headers:
                    all_unique_headers.append(h)

        suggestion, score = best_header_match(
            header,
            all_unique_headers
        )

        if not suggestion:
            return "No reliable match", score

        return suggestion, score

    def open_auto_header_popup(self):
        if not self.sheet_headers:
            messagebox.showwarning(
                "No Workbook",
                "Please select an Excel workbook first."
            )
            return

        HeaderMappingPopup(
            self.root,
            self.sheet_headers
        )

    # ========================================================
    # FILTER POPUP
    # ========================================================

    def open_filter_popup(self):
        if not self.sheet_data:
            messagebox.showwarning(
                "No Workbook",
                "Please select an Excel workbook first."
            )
            return

        popup = FilterPopup(
            self.root,
            self.sheet_data,
            initial_sheet=(
                next(iter(self.sheet_data.keys()))
                if self.sheet_data
                else None
            ),
            filtered_rows_by_sheet=self.filtered_rows_by_sheet
        )

        self.root.wait_window(popup)

        if popup.result:
            self.filtered_rows_by_sheet = (
                popup.result["filtered_rows_by_sheet"]
            )

            self.use_filtered_rows = (
                popup.result["use_filtered_rows"]
            )

            selected_sheet = popup.result["selected_sheet"]

            if selected_sheet:
                self.selected_sheet_var.set(
                    f"Selected sheet: {selected_sheet}"
                )

            if self.use_filtered_rows:
                total_filtered = sum(
                    len(df)
                    for df in self.filtered_rows_by_sheet.values()
                )

                self.filter_status_var.set(
                    f"FILTER MODE ON — {total_filtered:,} "
                    f"filtered row(s) will be used."
                )
            else:
                self.filter_status_var.set(
                    "FILTER MODE OFF — all source rows will be used."
                )

            self.status_var.set(
                "Sheet/filter selection updated."
            )

    # ========================================================
    # POPULATE TREEVIEW
    # ========================================================

    def populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.tree_metadata.clear()

        for sheet_name, headers in self.sheet_headers.items():
            for header_info in headers:
                suggestion, score = (
                    self.get_auto_suggestion(
                        header_info["original_header"]
                    )
                )

                item_id = self.tree.insert(
                    "",
                    "end",
                    values=(
                        sheet_name,
                        self.column_letter(
                            header_info["column_index"]
                        ),
                        header_info["header"],
                        suggestion,
                        f"{score:.1f}%"
                    )
                )

                self.tree_metadata[item_id] = (
                    sheet_name,
                    header_info
                )

                self.update_tree_item_color(
                    item_id,
                    header_info["color"]
                )

    # ========================================================
    # COLUMN NUMBER TO EXCEL LETTER
    # ========================================================

    def column_letter(self, number):
        result = ""
        number += 1

        while number:
            number, remainder = divmod(
                number - 1,
                26
            )

            result = (
                chr(65 + remainder)
                + result
            )

        return result

    # ========================================================
    # HEADER SELECTED
    # ========================================================

    def on_header_select(self, event=None):
        selected = self.tree.selection()

        if not selected:
            self.selected_item = None
            self.selected_header_label.config(text="None")
            return

        item_id = selected[0]
        self.selected_item = item_id

        if item_id not in self.tree_metadata:
            return

        sheet_name, header_info = (
            self.tree_metadata[item_id]
        )

        self.selected_header_label.config(
            text=(
                f"{sheet_name} → "
                f"{header_info['header']}"
            )
        )

        self.color_var.set(
            header_info["color"]
        )

    # ========================================================
    # APPLY COLOR
    # ========================================================

    def apply_color(self, event=None):
        if not self.selected_item:
            messagebox.showwarning(
                "No Header Selected",
                "Please select a header first."
            )
            return

        color = self.color_var.get()

        if not color:
            color = "None"

        sheet_name, header_info = (
            self.tree_metadata[
                self.selected_item
            ]
        )

        header_info["color"] = color

        values = list(
            self.tree.item(
                self.selected_item,
                "values"
            )
        )

        values[5] = color

        self.tree.item(
            self.selected_item,
            values=values
        )

        self.update_tree_item_color(
            self.selected_item,
            color
        )

        self.status_var.set(
            f"Color '{color}' assigned to "
            f"{sheet_name} → "
            f"{header_info['header']}"
        )

    # ========================================================
    # UPDATE TREEVIEW COLOR
    # ========================================================

    def update_tree_item_color(self, item_id, color):
        tag_name = f"color_{color}"

        if not self.tree.tag_has(tag_name):
            background = COLORS.get(
                color,
                "#FFFFFF"
            )

            self.tree.tag_configure(
                tag_name,
                background=background
            )

        self.tree.item(
            item_id,
            tags=(tag_name,)
        )

    # ========================================================
    # CLEAR ALL MAPPING
    # ========================================================

    def clear_mapping(self):
        if not self.sheet_headers:
            return

        answer = messagebox.askyesno(
            "Clear Mapping",
            "Are you sure you want to clear all "
            "header color mappings?"
        )

        if not answer:
            return

        for headers in self.sheet_headers.values():
            for header_info in headers:
                header_info["color"] = "None"

        self.populate_tree()

        self.color_var.set("None")
        self.selected_header_label.config(text="None")

        self.status_var.set(
            "All color mappings cleared."
        )

    # ========================================================
    # FIND COLOR CONFLICTS
    # ========================================================

    def get_color_conflicts(self):
        color_to_headers = OrderedDict()

        for sheet_name, headers in self.sheet_headers.items():
            for header_info in headers:
                color = header_info["color"]

                if color == "None":
                    continue

                header = header_info["original_header"]

                if color not in color_to_headers:
                    color_to_headers[color] = []

                if header not in color_to_headers[color]:
                    color_to_headers[color].append(header)

        conflicts = OrderedDict()

        for color, headers in color_to_headers.items():
            if len(headers) > 1:
                conflicts[color] = headers

        return conflicts

    # ========================================================
    # VALIDATE MAPPING
    # ========================================================

    def validate_mapping(self):
        if not self.sheet_headers:
            messagebox.showwarning(
                "No Workbook",
                "Please select an Excel workbook first."
            )
            return False

        conflicts = self.get_color_conflicts()

        if conflicts:
            conflict_text = "\n".join(
                f"• {color}: {', '.join(headers)}"
                for color, headers in conflicts.items()
            )

            messagebox.showwarning(
                "Color Mapping Warning",
                "The same color is assigned to different "
                "header names.\n\n"
                f"{conflict_text}\n\n"
                "This is only a warning.\n"
                "You can still use FINAL CONSOLIDATE."
            )

            self.status_var.set(
                "Warning: same color used for different headers."
            )

            return True

        messagebox.showinfo(
            "Mapping Valid",
            "Header color mapping is valid.\n\n"
            "No conflicting color mapping was found."
        )

        self.status_var.set(
            "Mapping validation successful."
        )

        return True

    # ========================================================
    # BUILD COLOR GROUPS
    # ========================================================

    def build_color_groups(self):
        color_groups = OrderedDict()

        for sheet_name, headers in self.sheet_headers.items():
            for header_info in headers:
                color = header_info["color"]

                if color == "None":
                    continue

                header = header_info["original_header"]

                if color not in color_groups:
                    color_groups[color] = []

                if header not in color_groups[color]:
                    color_groups[color].append(header)

        return color_groups

    # ========================================================
    # BUILD OUTPUT HEADERS
    # ========================================================

    def build_output_headers(self, color_groups):
        # Source Sheet is always first for reference.
        output_headers = ["Source Sheet"]

        # Colored groups.
        for color, headers in color_groups.items():
            if headers:
                if headers[0] not in output_headers:
                    output_headers.append(headers[0])

        # Uncolored headers.
        for sheet_name, headers in self.sheet_headers.items():
            for header_info in headers:
                if header_info["color"] != "None":
                    continue

                header = header_info["original_header"]

                if header not in output_headers:
                    output_headers.append(header)

        return output_headers

    # ========================================================
    # GET OUTPUT HEADER FOR SOURCE HEADER
    # ========================================================

    def get_output_header(self, header_info, color_groups):
        color = header_info["color"]
        original_header = header_info["original_header"]

        if color != "None" and color in color_groups:
            return color_groups[color][0]

        return original_header

    # ========================================================
    # GET DATA TO CONSOLIDATE
    # ========================================================

    def get_sheet_rows_for_consolidation(self, sheet_name):
        full_df = self.sheet_data[sheet_name]

        if not self.use_filtered_rows:
            return full_df

        # If a sheet has not been filtered, keep all its rows.
        if sheet_name not in self.filtered_rows_by_sheet:
            return full_df

        return self.filtered_rows_by_sheet[sheet_name]

    # ========================================================
    # FINAL CONSOLIDATION
    # ========================================================

    def final_consolidate(self):
        if not self.file_path:
            messagebox.showwarning(
                "No Excel File",
                "Please select an Excel workbook first."
            )
            return

        if not self.sheet_headers:
            messagebox.showwarning(
                "No Data",
                "No worksheets were loaded."
            )
            return

        conflicts = self.get_color_conflicts()

        if conflicts:
            conflict_text = "\n".join(
                f"• {color}: {', '.join(headers)}"
                for color, headers in conflicts.items()
            )

            answer = messagebox.askyesno(
                "Color Mapping Warning",
                "WARNING\n\n"
                "The same color has been assigned to different "
                "header names.\n\n"
                f"{conflict_text}\n\n"
                "The program will treat headers with the same "
                "color as the same consolidated field.\n\n"
                "Do you still want to continue with "
                "FINAL CONSOLIDATE?"
            )

            if not answer:
                self.status_var.set(
                    "Consolidation cancelled."
                )
                return

        filter_message = (
            "• Use the selected filtered rows\n"
            if self.use_filtered_rows
            else "• Use all source rows\n"
        )

        answer = messagebox.askyesno(
            "Final Consolidation",
            "Ready to consolidate the workbook.\n\n"
            "The program will:\n\n"
            "• Preserve every selected source row\n"
            f"{filter_message}"
            "• Add Source Sheet to every row\n"
            "• Align columns using header/color mapping\n"
            "• NOT match people or records\n"
            "• NOT deduplicate rows\n"
            "• NOT merge rows\n\n"
            "Do you want to continue?"
        )

        if not answer:
            return

        input_name = os.path.splitext(
            os.path.basename(self.file_path)
        )[0]

        default_name = (
            f"{input_name}_Consolidated.xlsx"
        )

        output_path = filedialog.asksaveasfilename(
            title="Save Consolidated Excel File",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[
                ("Excel Workbook", "*.xlsx")
            ]
        )

        if not output_path:
            return

        self.output_path = output_path

        try:
            self.status_var.set(
                "Consolidating data..."
            )
            self.root.update_idletasks()

            all_output_rows = []

            color_groups = self.build_color_groups()
            output_headers = self.build_output_headers(
                color_groups
            )

            total_rows = 0

            # Process every worksheet.
            for sheet_name in self.sheet_headers.keys():
                self.status_var.set(
                    f"Processing sheet: {sheet_name}"
                )
                self.root.update_idletasks()

                df = self.get_sheet_rows_for_consolidation(
                    sheet_name
                )

                if df is None or len(df.columns) == 0:
                    continue

                sheet_header_info = self.sheet_headers.get(
                    sheet_name,
                    []
                )

                column_output_map = {}

                for info in sheet_header_info:
                    column_index = info["column_index"]

                    if column_index >= len(df.columns):
                        continue

                    output_header = self.get_output_header(
                        info,
                        color_groups
                    )

                    column_output_map[
                        column_index
                    ] = output_header

                # Every source row remains a separate output row.
                for _, source_row in df.iterrows():
                    output_row = {
                        header: ""
                        for header in output_headers
                    }

                    # Source Sheet reference.
                    output_row["Source Sheet"] = sheet_name

                    has_data = False

                    for (
                        column_index,
                        output_header
                    ) in column_output_map.items():

                        if column_index >= len(source_row):
                            continue

                        value = source_row.iloc[column_index]

                        if pd.notna(value):
                            output_row[output_header] = value
                            has_data = True

                    # Keep the row if it contains data.
                    if has_data:
                        all_output_rows.append(
                            output_row
                        )
                        total_rows += 1

            result_df = pd.DataFrame(
                all_output_rows,
                columns=output_headers
            )

            # Save output.
            result_df.to_excel(
                output_path,
                index=False,
                sheet_name="Consolidated"
            )

            self.status_var.set(
                f"Completed: {total_rows:,} rows consolidated."
            )

            messagebox.showinfo(
                "Consolidation Complete",
                "Consolidation completed successfully.\n\n"
                f"Sheets processed: "
                f"{len(self.sheet_headers)}\n"
                f"Output columns: "
                f"{len(output_headers)}\n"
                f"Output rows: "
                f"{total_rows:,}\n\n"
                f"Source Sheet column included.\n\n"
                f"Saved to:\n"
                f"{output_path}"
            )

            open_folder = messagebox.askyesno(
                "Open Output Folder",
                "Do you want to open the output folder?"
            )

            if open_folder:
                folder = os.path.dirname(output_path)

                try:
                    os.startfile(folder)
                except Exception:
                    pass

        except PermissionError:
            messagebox.showerror(
                "Permission Error",
                "Unable to save the output file.\n\n"
                "Please make sure the output Excel file is not "
                "already open."
            )

            self.status_var.set(
                "Unable to save output file."
            )

        except Exception as e:
            messagebox.showerror(
                "Consolidation Error",
                "An error occurred during consolidation.\n\n"
                f"{type(e).__name__}: {e}"
            )

            self.status_var.set(
                "Consolidation failed."
            )


# ============================================================
# MAIN
# ============================================================

def main():
    root = tk.Tk()
    app = HeaderConsolidatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
