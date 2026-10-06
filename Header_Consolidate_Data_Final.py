import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re
from collections import OrderedDict
from difflib import SequenceMatcher

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

# Common header aliases. These are used only for HEADER correspondence,
# never for person/record matching.
HEADER_ALIASES = {
    "document_id": {
        "document id", "document no", "document number", "doc id",
        "doc no", "doc number", "documentid", "documentno", "documentnumber"
    },
    "name": {
        "name", "full name", "customer name", "client name",
        "person name", "applicant name", "member name"
    },
    "tin": {
        "tin", "tin number", "tax number", "tax no", "tax id",
        "tax identification number", "tin no", "tin id"
    },
    "ssn": {
        "ssn", "ssn number", "ssn no", "social security number"
    },
    "dl": {
        "dl", "dl number", "dl no", "driver license",
        "driver license number", "driving license", "driving licence",
        "drivers license", "drivers license number"
    },
    "dob": {
        "dob", "date of birth", "birth date", "birthdate"
    },
    "provider_id": {
        "provider id", "provider no", "provider number", "providerid"
    },
}

AUTO_COLORS = [
    "Yellow", "Blue", "Green", "Orange", "Pink",
    "Purple", "Red", "Light Blue", "Light Green", "Grey"
]


def normalize_header(value):
    """Normalize a header for semantic comparison."""
    text = str(value).strip().lower()
    text = text.replace("&", " and ")
    text = re.sub(r"[_\-./]+", " ", text)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def compact_header(value):
    return re.sub(r"[^a-z0-9]", "", normalize_header(value))


def alias_key(header):
    norm = normalize_header(header)
    compact = compact_header(header)

    for key, aliases in HEADER_ALIASES.items():
        if norm in aliases or compact in {compact_header(x) for x in aliases}:
            return key
    return None


def header_similarity(a, b):
    na = normalize_header(a)
    nb = normalize_header(b)
    if not na or not nb:
        return 0.0

    ka = alias_key(a)
    kb = alias_key(b)

    # Strong semantic match when both headers belong to the same known group.
    if ka and kb and ka == kb:
        return 1.0

    # Never treat different known semantic groups as a match.
    if ka and kb and ka != kb:
        return 0.0

    ca = compact_header(a)
    cb = compact_header(b)
    return SequenceMatcher(None, ca, cb).ratio()


class HeaderConsolidatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1450x850")
        self.root.minsize(1100, 700)

        self.file_path = None
        self.output_path = None
        self.sheet_headers = OrderedDict()
        self.tree_metadata = {}
        self.selected_item = None

        self.sheet_filter_var = tk.StringVar(value="All Sheets")
        self.header_filter_var = tk.StringVar(value="")
        self.color_filter_var = tk.StringVar(value="All Colors")
        self.color_var = tk.StringVar(value="None")

        self.create_styles()
        self.create_ui()

    # ------------------------------------------------------------
    # UI
    # ------------------------------------------------------------
    def create_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Action.TButton", font=("Segoe UI", 10, "bold"), padding=8)

    def create_ui(self):
        top = ttk.Frame(self.root, padding=15)
        top.pack(fill="x")

        ttk.Label(
            top, text=APP_TITLE, style="Title.TLabel"
        ).pack(anchor="w")

        ttk.Label(
            top,
            text=(
                "Map corresponding Excel headers using colors, keep every source row, "
                "and add the originating Sheet Name for reference."
            ),
            style="Subtitle.TLabel"
        ).pack(anchor="w", pady=(3, 10))

        file_frame = ttk.LabelFrame(
            self.root, text="1. Select Excel File", padding=12
        )
        file_frame.pack(fill="x", padx=15, pady=(0, 8))

        self.file_label = ttk.Label(
            file_frame, text="No Excel file selected.", foreground="gray"
        )
        self.file_label.pack(side="left", fill="x", expand=True)

        ttk.Button(
            file_frame, text="Select Excel File",
            command=self.select_file, style="Action.TButton"
        ).pack(side="right")

        mapping = ttk.LabelFrame(
            self.root, text="2. Header Mapping / Filtering", padding=10
        )
        mapping.pack(fill="both", expand=True, padx=15, pady=(0, 8))

        instruction = (
            "Use the filters like an Excel-style view to find headers quickly. "
            "You can manually assign colors or use AUTO IDENTIFY HEADERS. "
            "Same-color headers become one output field. Color is only a header "
            "correspondence key — never a record/person matching key."
        )
        ttk.Label(mapping, text=instruction, justify="left").pack(
            anchor="w", pady=(0, 8)
        )

        filter_frame = ttk.Frame(mapping)
        filter_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(filter_frame, text="Sheet:").pack(side="left")
        self.sheet_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.sheet_filter_var,
            state="readonly",
            width=28
        )
        self.sheet_combo.pack(side="left", padx=(5, 15))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        ttk.Label(filter_frame, text="Header Search:").pack(side="left")
        self.header_entry = ttk.Entry(
            filter_frame, textvariable=self.header_filter_var, width=30
        )
        self.header_entry.pack(side="left", padx=(5, 15))
        self.header_entry.bind("<KeyRelease>", lambda e: self.apply_filters())

        ttk.Label(filter_frame, text="Color:").pack(side="left")
        self.color_filter_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.color_filter_var,
            values=["All Colors"] + list(COLORS.keys()),
            state="readonly",
            width=18
        )
        self.color_filter_combo.pack(side="left", padx=5)
        self.color_filter_combo.bind(
            "<<ComboboxSelected>>", lambda e: self.apply_filters()
        )

        ttk.Button(
            filter_frame, text="Clear Filters",
            command=self.clear_filters
        ).pack(side="left", padx=8)

        ttk.Button(
            filter_frame, text="AUTO IDENTIFY HEADERS",
            command=self.auto_identify_headers,
            style="Action.TButton"
        ).pack(side="right")

        table_frame = ttk.Frame(mapping)
        table_frame.pack(fill="both", expand=True)

        columns = ("sheet", "column", "header", "suggestion", "color")
        self.tree = ttk.Treeview(
            table_frame, columns=columns, show="headings", selectmode="browse"
        )

        headings = {
            "sheet": "Sheet",
            "column": "Column",
            "header": "Header Name",
            "suggestion": "Automatic Match / Suggestion",
            "color": "Assigned Color",
        }
        widths = {
            "sheet": 180,
            "column": 80,
            "header": 280,
            "suggestion": 330,
            "color": 160,
        }

        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w")

        scrollbar_y = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.tree.yview
        )
        scrollbar_x = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.tree.xview
        )
        self.tree.configure(
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self.on_header_select)

        color_frame = ttk.Frame(mapping)
        color_frame.pack(fill="x", pady=(8, 0))

        ttk.Label(color_frame, text="Selected Header:").pack(side="left")
        self.selected_header_label = ttk.Label(
            color_frame, text="None", font=("Segoe UI", 10, "bold")
        )
        self.selected_header_label.pack(side="left", padx=(5, 20))

        ttk.Label(color_frame, text="Color:").pack(side="left")
        self.color_combo = ttk.Combobox(
            color_frame,
            textvariable=self.color_var,
            values=list(COLORS.keys()),
            state="readonly",
            width=18
        )
        self.color_combo.pack(side="left", padx=8)
        self.color_combo.bind("<<ComboboxSelected>>", self.apply_color)

        ttk.Button(
            color_frame, text="Apply Color",
            command=self.apply_color
        ).pack(side="left")

        bottom = ttk.Frame(self.root, padding=(15, 0, 15, 15))
        bottom.pack(fill="x")

        ttk.Button(
            bottom, text="Validate Mapping",
            command=self.validate_mapping,
            style="Action.TButton"
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            bottom, text="Clear Mapping",
            command=self.clear_mapping
        ).pack(side="left")

        ttk.Button(
            bottom, text="FINAL CONSOLIDATE",
            command=self.final_consolidate,
            style="Action.TButton"
        ).pack(side="right")

        self.status_var = tk.StringVar(value="Please select an Excel file.")
        ttk.Label(
            self.root, textvariable=self.status_var,
            relief="sunken", anchor="w", padding=5
        ).pack(side="bottom", fill="x")

    # ------------------------------------------------------------
    # File loading
    # ------------------------------------------------------------
    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel Workbook",
            filetypes=[
                ("Excel Files", "*.xlsx *.xlsm"),
                ("Excel Workbook", "*.xlsx"),
                ("Excel Macro Workbook", "*.xlsm"),
                ("All Files", "*.*"),
            ]
        )
        if not path:
            return

        self.file_path = path
        self.file_label.config(text=os.path.basename(path), foreground="black")
        self.load_workbook_headers()

    def load_workbook_headers(self):
        try:
            self.status_var.set("Reading workbook...")
            self.root.update_idletasks()

            excel = pd.ExcelFile(self.file_path)
            self.sheet_headers.clear()

            for sheet_name in excel.sheet_names:
                df = pd.read_excel(
                    self.file_path, sheet_name=sheet_name, nrows=0
                )

                headers = []
                used_names = {}

                for column_index, value in enumerate(df.columns):
                    header = str(value).strip()
                    if not header or header.lower().startswith("unnamed:"):
                        header = f"Unnamed Column {column_index + 1}"

                    if header in used_names:
                        used_names[header] += 1
                        display_header = (
                            f"{header} (Duplicate {used_names[header]})"
                        )
                    else:
                        used_names[header] = 1
                        display_header = header

                    headers.append({
                        "column_index": column_index,
                        "header": display_header,
                        "original_header": header,
                        "color": "None",
                        "suggestion": self.get_semantic_suggestion(header),
                    })

                self.sheet_headers[sheet_name] = headers

            self.update_sheet_filter_values()
            self.populate_tree()

            total_headers = sum(
                len(headers) for headers in self.sheet_headers.values()
            )
            self.status_var.set(
                f"Loaded {len(self.sheet_headers)} sheet(s) and "
                f"{total_headers} header(s)."
            )

            if not self.sheet_headers:
                messagebox.showwarning(
                    "No Sheets", "The workbook does not contain any worksheets."
                )

        except Exception as e:
            messagebox.showerror(
                "Error Reading Excel",
                f"Unable to read the workbook.\n\n{type(e).__name__}: {e}"
            )
            self.status_var.set("Error reading workbook.")

    # ------------------------------------------------------------
    # Header suggestions / automatic identification
    # ------------------------------------------------------------
    def get_semantic_suggestion(self, header):
        key = alias_key(header)
        if key:
            pretty = {
                "document_id": "Document ID",
                "name": "Name",
                "tin": "TIN",
                "ssn": "SSN",
                "dl": "DL",
                "dob": "DOB",
                "provider_id": "Provider ID",
            }
            return pretty.get(key, key.replace("_", " ").title())

        # Find a close header from the workbook after it is loaded.
        candidates = []
        for _, headers in self.sheet_headers.items():
            for info in headers:
                other = info["original_header"]
                if other == header:
                    continue
                score = header_similarity(header, other)
                if score >= 0.78:
                    candidates.append((score, other))

        if candidates:
            candidates.sort(reverse=True)
            return candidates[0][1]

        return "No automatic match"

    def auto_identify_headers(self):
        if not self.sheet_headers:
            messagebox.showwarning(
                "No Workbook", "Please select an Excel workbook first."
            )
            return

        # Recalculate suggestions using all loaded headers.
        all_headers = []
        for _, headers in self.sheet_headers.items():
            for info in headers:
                all_headers.append(info["original_header"])

        unique_headers = list(OrderedDict.fromkeys(all_headers))
        groups = []
        assigned = {}

        # First use known semantic aliases.
        for header in unique_headers:
            key = alias_key(header)
            if key:
                assigned[header] = key

        # Then use conservative fuzzy grouping for otherwise-unmapped headers.
        for header in unique_headers:
            if header in assigned:
                continue

            best_key = None
            best_score = 0.0

            for existing_header, key in assigned.items():
                score = header_similarity(header, existing_header)
                if score > best_score:
                    best_score = score
                    best_key = key

            if best_key and best_score >= 0.88:
                assigned[header] = best_key
            else:
                assigned[header] = f"custom_{len(groups)}_{header}"

            if assigned[header].startswith("custom_"):
                groups.append(assigned[header])

        # Stable group order based on first appearance.
        group_order = []
        for header in unique_headers:
            key = assigned[header]
            if key not in group_order:
                group_order.append(key)

        color_for_group = {}
        color_index = 0

        for key in group_order:
            # If more groups than available colors, leave additional groups None.
            if color_index < len(AUTO_COLORS):
                color_for_group[key] = AUTO_COLORS[color_index]
                color_index += 1
            else:
                color_for_group[key] = "None"

        changed = 0
        for _, headers in self.sheet_headers.items():
            for info in headers:
                key = assigned.get(info["original_header"])
                if key is None:
                    continue
                new_color = color_for_group.get(key, "None")
                if info["color"] != new_color:
                    info["color"] = new_color
                    changed += 1
                info["suggestion"] = self.get_group_display_name(
                    info["original_header"], key
                )

        self.populate_tree()
        self.status_var.set(
            f"Automatic identification completed. {changed} header mapping(s) updated."
        )

        messagebox.showinfo(
            "Automatic Header Identification",
            "Automatic header identification completed.\n\n"
            "Please review the colors before FINAL CONSOLIDATE. "
            "You can manually change any mapping."
        )

    def get_group_display_name(self, header, key):
        if key in HEADER_ALIASES:
            return key.replace("_", " ").title()
        return "Possible Match: " + str(header)

    # ------------------------------------------------------------
    # Filtering
    # ------------------------------------------------------------
    def update_sheet_filter_values(self):
        values = ["All Sheets"] + list(self.sheet_headers.keys())
        self.sheet_combo["values"] = values
        self.sheet_filter_var.set("All Sheets")

    def clear_filters(self):
        self.sheet_filter_var.set("All Sheets")
        self.header_filter_var.set("")
        self.color_filter_var.set("All Colors")
        self.populate_tree()

    def apply_filters(self):
        sheet_filter = self.sheet_filter_var.get()
        header_filter = normalize_header(self.header_filter_var.get())
        color_filter = self.color_filter_var.get()

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.tree_metadata.clear()

        for sheet_name, headers in self.sheet_headers.items():
            if sheet_filter != "All Sheets" and sheet_name != sheet_filter:
                continue

            for info in headers:
                header_norm = normalize_header(info["header"])
                if header_filter and header_filter not in header_norm:
                    continue

                if color_filter != "All Colors" and info["color"] != color_filter:
                    continue

                self.insert_tree_item(sheet_name, info)

    # ------------------------------------------------------------
    # Treeview
    # ------------------------------------------------------------
    def populate_tree(self):
        self.apply_filters()

    def insert_tree_item(self, sheet_name, header_info):
        item_id = self.tree.insert(
            "",
            "end",
            values=(
                sheet_name,
                self.column_letter(header_info["column_index"]),
                header_info["header"],
                header_info.get("suggestion", "No automatic match"),
                header_info["color"],
            )
        )

        self.tree_metadata[item_id] = (sheet_name, header_info)
        self.update_tree_item_color(item_id, header_info["color"])

    def column_letter(self, number):
        result = ""
        number += 1
        while number:
            number, remainder = divmod(number - 1, 26)
            result = chr(65 + remainder) + result
        return result

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

        sheet_name, header_info = self.tree_metadata[item_id]
        self.selected_header_label.config(
            text=f"{sheet_name} → {header_info['header']}"
        )
        self.color_var.set(header_info["color"])

    def apply_color(self, event=None):
        if not self.selected_item:
            messagebox.showwarning(
                "No Header Selected", "Please select a header first."
            )
            return

        color = self.color_var.get() or "None"
        sheet_name, header_info = self.tree_metadata[self.selected_item]
        header_info["color"] = color

        values = list(self.tree.item(self.selected_item, "values"))
        values[4] = color
        self.tree.item(self.selected_item, values=values)
        self.update_tree_item_color(self.selected_item, color)

        self.status_var.set(
            f"Color '{color}' assigned to {sheet_name} → {header_info['header']}"
        )

    def update_tree_item_color(self, item_id, color):
        tag_name = f"color_{color.replace(' ', '_')}"
        if not self.tree.tag_has(tag_name):
            self.tree.tag_configure(
                tag_name, background=COLORS.get(color, "#FFFFFF")
            )
        self.tree.item(item_id, tags=(tag_name,))

    # ------------------------------------------------------------
    # Mapping validation
    # ------------------------------------------------------------
    def clear_mapping(self):
        if not self.sheet_headers:
            return

        if not messagebox.askyesno(
            "Clear Mapping",
            "Are you sure you want to clear all header color mappings?"
        ):
            return

        for headers in self.sheet_headers.values():
            for info in headers:
                info["color"] = "None"

        self.color_var.set("None")
        self.selected_header_label.config(text="None")
        self.populate_tree()
        self.status_var.set("All color mappings cleared.")

    def get_color_conflicts(self):
        color_to_headers = OrderedDict()

        for _, headers in self.sheet_headers.items():
            for info in headers:
                color = info["color"]
                if color == "None":
                    continue

                header = info["original_header"]
                color_to_headers.setdefault(color, [])

                if header not in color_to_headers[color]:
                    color_to_headers[color].append(header)

        return OrderedDict(
            (color, headers)
            for color, headers in color_to_headers.items()
            if len(headers) > 1
        )

    def validate_mapping(self):
        if not self.sheet_headers:
            messagebox.showwarning(
                "No Workbook", "Please select an Excel workbook first."
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
                "The same color is assigned to different header names.\n\n"
                f"{conflict_text}\n\n"
                "This is only a warning. You can still use FINAL CONSOLIDATE."
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
        self.status_var.set("Mapping validation successful.")
        return True

    # ------------------------------------------------------------
    # Output mapping
    # ------------------------------------------------------------
    def build_color_groups(self):
        color_groups = OrderedDict()

        for _, headers in self.sheet_headers.items():
            for info in headers:
                color = info["color"]
                if color == "None":
                    continue

                header = info["original_header"]
                color_groups.setdefault(color, [])

                if header not in color_groups[color]:
                    color_groups[color].append(header)

        return color_groups

    def build_output_headers(self, color_groups):
        output_headers = []

        # First header found for a color becomes the output field.
        for _, headers in color_groups.items():
            if headers:
                output_headers.append(headers[0])

        # Uncolored headers are retained as their original names.
        for _, headers in self.sheet_headers.items():
            for info in headers:
                if info["color"] != "None":
                    continue

                header = info["original_header"]
                if header not in output_headers:
                    output_headers.append(header)

        return output_headers

    def get_output_header(self, header_info, color_groups):
        color = header_info["color"]
        original_header = header_info["original_header"]

        if color != "None" and color in color_groups:
            return color_groups[color][0]

        return original_header

    # ------------------------------------------------------------
    # Final consolidation
    # ------------------------------------------------------------
    def final_consolidate(self):
        if not self.file_path:
            messagebox.showwarning(
                "No Excel File",
                "Please select an Excel workbook first."
            )
            return

        if not self.sheet_headers:
            messagebox.showwarning(
                "No Data", "No worksheets were loaded."
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
                "The same color has been assigned to different header names.\n\n"
                f"{conflict_text}\n\n"
                "Headers with the same color will be treated as one "
                "consolidated field.\n\n"
                "Do you still want to continue with FINAL CONSOLIDATE?"
            )
            if not answer:
                self.status_var.set("Consolidation cancelled.")
                return

        answer = messagebox.askyesno(
            "Final Consolidation",
            "Ready to consolidate the workbook.\n\n"
            "The program will:\n\n"
            "• Preserve every source row\n"
            "• Append rows from all sheets\n"
            "• Add Source Sheet for reference\n"
            "• Align columns using header/color mapping\n"
            "• Automatically use the first header for each color group\n"
            "• Split very large exports into multiple Excel sheets\n"
            "• NOT match people or records\n"
            "• NOT deduplicate rows\n"
            "• NOT merge rows\n\n"
            "Do you want to continue?"
        )
        if not answer:
            return

        input_name = os.path.splitext(os.path.basename(self.file_path))[0]
        default_name = f"{input_name}_Consolidated.xlsx"

        output_path = filedialog.asksaveasfilename(
            title="Save Consolidated Excel File",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not output_path:
            return

        self.output_path = output_path

        try:
            self.status_var.set("Consolidating data...")
            self.root.update_idletasks()

            excel = pd.ExcelFile(self.file_path)
            color_groups = self.build_color_groups()
            output_headers = self.build_output_headers(color_groups)

            # Source Sheet is always included as the first reference column.
            final_headers = ["Source Sheet"] + output_headers

            all_output_rows = []
            total_rows = 0

            for sheet_name in excel.sheet_names:
                self.status_var.set(f"Processing sheet: {sheet_name}")
                self.root.update_idletasks()

                df = pd.read_excel(
                    self.file_path,
                    sheet_name=sheet_name,
                    dtype=object
                )

                if len(df.columns) == 0:
                    continue

                sheet_header_info = self.sheet_headers.get(sheet_name, [])

                column_output_map = {}
                for info in sheet_header_info:
                    idx = info["column_index"]
                    if idx >= len(df.columns):
                        continue

                    output_header = self.get_output_header(
                        info, color_groups
                    )
                    column_output_map[idx] = output_header

                for _, source_row in df.iterrows():
                    output_row = {header: "" for header in final_headers}
                    output_row["Source Sheet"] = sheet_name
                    has_data = False

                    for column_index, output_header in column_output_map.items():
                        if column_index >= len(source_row):
                            continue

                        value = source_row.iloc[column_index]
                        if pd.isna(value):
                            continue

                        # If two source columns map to the same output header
                        # within one row, preserve both values instead of
                        # silently overwriting one.
                        existing = output_row.get(output_header, "")
                        if existing not in ("", None) and str(existing) != str(value):
                            output_row[output_header] = (
                                f"{existing} | {value}"
                            )
                        else:
                            output_row[output_header] = value

                        has_data = True

                    if has_data:
                        all_output_rows.append(output_row)
                        total_rows += 1

            result_df = pd.DataFrame(
                all_output_rows,
                columns=final_headers
            )

            # ------------------------------------------------------------
            # LARGE DATA EXPORT
            # ------------------------------------------------------------
            # Excel supports a maximum of 1,048,576 rows per worksheet.
            # The first row is the header, so data is automatically divided
            # into safe-sized worksheet parts.
            MAX_DATA_ROWS_PER_SHEET = 1_048_575

            total_export_parts = max(
                1,
                (len(result_df) + MAX_DATA_ROWS_PER_SHEET - 1)
                // MAX_DATA_ROWS_PER_SHEET
            )

            self.status_var.set(
                f"Preparing export: {len(result_df):,} rows "
                f"across {total_export_parts} Excel sheet(s)..."
            )
            self.root.update_idletasks()

            # Export in parts so large data never exceeds Excel's worksheet
            # row limit. No rows are discarded.
            with pd.ExcelWriter(
                output_path, engine="openpyxl"
            ) as writer:

                if result_df.empty:
                    result_df.to_excel(
                        writer,
                        index=False,
                        sheet_name="Consolidated_Part_1"
                    )
                    parts_created = 1

                else:
                    parts_created = 0

                    for start_row in range(
                        0, len(result_df), MAX_DATA_ROWS_PER_SHEET
                    ):
                        end_row = min(
                            start_row + MAX_DATA_ROWS_PER_SHEET,
                            len(result_df)
                        )

                        part_number = parts_created + 1
                        sheet_name = f"Consolidated_Part_{part_number}"

                        self.status_var.set(
                            f"Exporting Part {part_number} of "
                            f"{total_export_parts}: rows "
                            f"{start_row + 1:,} - {end_row:,}"
                        )
                        self.root.update_idletasks()

                        part_df = result_df.iloc[start_row:end_row]

                        part_df.to_excel(
                            writer,
                            index=False,
                            sheet_name=sheet_name
                        )

                        ws = writer.book[sheet_name]
                        ws.freeze_panes = "A2"
                        ws.auto_filter.ref = ws.dimensions

                        # Reasonable column widths without allowing very long
                        # cell values to make the workbook unnecessarily wide.
                        for column_cells in ws.columns:
                            max_length = 0
                            column_letter = column_cells[0].column_letter

                            for cell in column_cells:
                                try:
                                    value_length = len(str(cell.value))
                                    max_length = max(
                                        max_length,
                                        min(value_length, 45)
                                    )
                                except Exception:
                                    pass

                            ws.column_dimensions[column_letter].width = min(
                                max(max_length + 2, 12), 45
                            )

                        parts_created += 1

            self.status_var.set(
                f"Completed: {total_rows:,} rows exported into "
                f"{parts_created} Excel sheet(s)."
            )

            messagebox.showinfo(
                "Consolidation Complete",
                "Consolidation and export completed successfully.\n\n"
                f"Sheets processed: {len(excel.sheet_names)}\n"
                f"Output columns: {len(final_headers)}\n"
                f"Output rows: {total_rows:,}\n"
                f"Excel output parts: {parts_created}\n\n"
                "Large data was automatically divided into multiple "
                "Excel sheets to stay within Excel's row limit.\n\n"
                "Every exported row is preserved.\n"
                "Source Sheet column is included for reference.\n\n"
                f"Saved to:\n{output_path}"
            )

            if messagebox.askyesno(
                "Open Output Folder",
                "Do you want to open the output folder?"
            ):
                folder = os.path.dirname(output_path)
                try:
                    os.startfile(folder)
                except Exception:
                    pass

        except PermissionError:
            messagebox.showerror(
                "Permission Error",
                "Unable to save the output file.\n\n"
                "Please make sure the output Excel file is not already open."
            )
            self.status_var.set("Unable to save output file.")

        except Exception as e:
            messagebox.showerror(
                "Consolidation Error",
                f"An error occurred during consolidation.\n\n"
                f"{type(e).__name__}: {e}"
            )
            self.status_var.set("Consolidation failed.")


def main():
    root = tk.Tk()
    app = HeaderConsolidatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
