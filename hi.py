import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from collections import OrderedDict


# ============================================================
# HEADER CONSOLIDATE DATA
# ============================================================
#
# FINAL LOGIC
#
# 1. Select an Excel workbook.
# 2. Read all worksheets.
# 3. Read all headers.
# 4. Assign colors to corresponding headers.
#
# Example:
#
# Sheet 1:
#     Document ID -> Yellow
#     Name        -> Blue
#     TIN         -> Green
#
# Sheet 2:
#     Document No -> Yellow
#     Full Name   -> Blue
#     Tax Number  -> Green
#
# Result:
#
#     Document ID | Name | TIN
#
# IMPORTANT:
#
# - Color is used for HEADER CORRESPONDENCE.
# - Color is NOT used for record matching.
# - Color is NOT used for identifying people.
# - No Document ID matching.
# - No Name matching.
# - No TIN matching.
# - No deduplication.
# - No row merging.
# - Every source row is preserved.
#
# IMPORTANT NEW RULE:
#
# If the same color is assigned to different header names:
#
#     Document ID     -> Yellow
#     Document Number -> Yellow
#
# The program shows a WARNING.
#
# BUT:
#
# The user can still continue with FINAL CONSOLIDATE.
#
# If the user clicks YES, the consolidation proceeds.
#
# The first header name found for that color becomes
# the consolidated output header.
#
# ============================================================


APP_TITLE = "Header Consolidate Data"


# ============================================================
# AVAILABLE COLORS
# ============================================================

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


class HeaderConsolidatorApp:

    # ========================================================
    # INITIALIZATION
    # ========================================================

    def __init__(self, root):

        self.root = root

        self.root.title(APP_TITLE)
        self.root.geometry("1200x750")
        self.root.minsize(1000, 650)

        # Selected Excel file
        self.file_path = None

        # Output file
        self.output_path = None

        # ----------------------------------------------------
        # Header information
        #
        # {
        #   sheet_name: [
        #       {
        #           "column_index": 0,
        #           "header": "Document ID",
        #           "original_header": "Document ID",
        #           "color": "Yellow"
        #       }
        #   ]
        # }
        # ----------------------------------------------------

        self.sheet_headers = OrderedDict()

        # Treeview item metadata
        self.tree_metadata = {}

        # Currently selected Treeview item
        self.selected_item = None

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
    # CREATE USER INTERFACE
    # ========================================================

    def create_ui(self):

        # ----------------------------------------------------
        # TOP AREA
        # ----------------------------------------------------

        top_frame = ttk.Frame(
            self.root,
            padding=15
        )

        top_frame.pack(
            fill="x"
        )

        ttk.Label(
            top_frame,
            text="Header Consolidate Data",
            style="Title.TLabel"
        ).pack(
            anchor="w"
        )

        ttk.Label(
            top_frame,
            text=(
                "Map corresponding Excel headers using colors "
                "and consolidate all rows without record "
                "matching or deduplication."
            ),
            style="Subtitle.TLabel"
        ).pack(
            anchor="w",
            pady=(3, 10)
        )

        # ----------------------------------------------------
        # FILE SELECTION
        # ----------------------------------------------------

        file_frame = ttk.LabelFrame(
            self.root,
            text="1. Select Excel File",
            padding=12
        )

        file_frame.pack(
            fill="x",
            padx=15,
            pady=(0, 10)
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
        ).pack(
            side="right"
        )

        # ----------------------------------------------------
        # HEADER MAPPING AREA
        # ----------------------------------------------------

        instruction_frame = ttk.LabelFrame(
            self.root,
            text="2. Header Color Mapping",
            padding=10
        )

        instruction_frame.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=(0, 10)
        )

        instruction = (
            "Assign the SAME color to headers that represent "
            "the same logical field across sheets.\n\n"
            "Example: Document ID and Document Number can "
            "both be Yellow if they represent the same field.\n\n"
            "IMPORTANT: If the same color is assigned to "
            "different header names, the program will show "
            "a warning, but you can still continue with "
            "FINAL CONSOLIDATE."
        )

        ttk.Label(
            instruction_frame,
            text=instruction,
            justify="left"
        ).pack(
            anchor="w",
            pady=(0, 8)
        )

        # ----------------------------------------------------
        # TABLE FRAME
        # ----------------------------------------------------

        table_frame = ttk.Frame(
            instruction_frame
        )

        table_frame.pack(
            fill="both",
            expand=True
        )

        columns = (
            "sheet",
            "column",
            "header",
            "color"
        )

        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse"
        )

        # Headings
        self.tree.heading(
            "sheet",
            text="Sheet"
        )

        self.tree.heading(
            "column",
            text="Column"
        )

        self.tree.heading(
            "header",
            text="Header Name"
        )

        self.tree.heading(
            "color",
            text="Assigned Color"
        )

        # Column widths
        self.tree.column(
            "sheet",
            width=180
        )

        self.tree.column(
            "column",
            width=90
        )

        self.tree.column(
            "header",
            width=400
        )

        self.tree.column(
            "color",
            width=180
        )

        # Vertical scrollbar
        scrollbar_y = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )

        # Horizontal scrollbar
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

        table_frame.rowconfigure(
            0,
            weight=1
        )

        table_frame.columnconfigure(
            0,
            weight=1
        )

        self.tree.bind(
            "<<TreeviewSelect>>",
            self.on_header_select
        )

        # ----------------------------------------------------
        # COLOR CONTROLS
        # ----------------------------------------------------

        color_frame = ttk.Frame(
            instruction_frame
        )

        color_frame.pack(
            fill="x",
            pady=(10, 0)
        )

        ttk.Label(
            color_frame,
            text="Selected Header:"
        ).pack(
            side="left"
        )

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
        ).pack(
            side="left"
        )

        self.color_var = tk.StringVar(
            value="None"
        )

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
        ).pack(
            side="left"
        )

        # ----------------------------------------------------
        # BOTTOM BUTTONS
        # ----------------------------------------------------

        bottom_frame = ttk.Frame(
            self.root,
            padding=(15, 0, 15, 15)
        )

        bottom_frame.pack(
            fill="x"
        )

        ttk.Button(
            bottom_frame,
            text="Validate Mapping",
            command=self.validate_mapping,
            style="Action.TButton"
        ).pack(
            side="left",
            padx=(0, 8)
        )

        ttk.Button(
            bottom_frame,
            text="Clear Mapping",
            command=self.clear_mapping
        ).pack(
            side="left"
        )

        ttk.Button(
            bottom_frame,
            text="FINAL CONSOLIDATE",
            command=self.final_consolidate,
            style="Action.TButton"
        ).pack(
            side="right"
        )

        # ----------------------------------------------------
        # STATUS BAR
        # ----------------------------------------------------

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
                (
                    "Excel Files",
                    "*.xlsx *.xlsm"
                ),
                (
                    "Excel Workbook",
                    "*.xlsx"
                ),
                (
                    "Excel Macro Workbook",
                    "*.xlsm"
                ),
                (
                    "All Files",
                    "*.*"
                )
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
    # LOAD WORKBOOK HEADERS
    # ========================================================

    def load_workbook_headers(self):

        try:

            self.status_var.set(
                "Reading workbook..."
            )

            self.root.update_idletasks()

            excel = pd.ExcelFile(
                self.file_path
            )

            self.sheet_headers.clear()

            # ------------------------------------------------
            # Read every worksheet
            # ------------------------------------------------

            for sheet_name in excel.sheet_names:

                df = pd.read_excel(
                    self.file_path,
                    sheet_name=sheet_name,
                    nrows=0
                )

                headers = []

                used_names = {}

                # ------------------------------------------------
                # Read headers
                # ------------------------------------------------

                for column_index, value in enumerate(
                    df.columns
                ):

                    header = str(value).strip()

                    if not header:

                        header = (
                            f"Unnamed Column "
                            f"{column_index + 1}"
                        )

                    # --------------------------------------------
                    # Handle duplicate header names in SAME sheet
                    # --------------------------------------------

                    if header in used_names:

                        used_names[header] += 1

                        display_header = (
                            f"{header} "
                            f"(Duplicate "
                            f"{used_names[header]})"
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

                self.sheet_headers[
                    sheet_name
                ] = headers

            # ------------------------------------------------
            # Populate GUI
            # ------------------------------------------------

            self.populate_tree()

            total_headers = sum(
                len(headers)
                for headers in (
                    self.sheet_headers.values()
                )
            )

            self.status_var.set(
                f"Loaded "
                f"{len(self.sheet_headers)} sheet(s) "
                f"and "
                f"{total_headers} header(s)."
            )

            if not self.sheet_headers:

                messagebox.showwarning(
                    "No Sheets",
                    "The workbook does not contain "
                    "any worksheets."
                )

        except Exception as e:

            messagebox.showerror(
                "Error Reading Excel",
                (
                    "Unable to read the workbook.\n\n"
                    f"{type(e).__name__}: {e}"
                )
            )

            self.status_var.set(
                "Error reading workbook."
            )

    # ========================================================
    # POPULATE TREEVIEW
    # ========================================================

    def populate_tree(self):

        # Remove old items
        for item in self.tree.get_children():

            self.tree.delete(
                item
            )

        self.tree_metadata.clear()

        # ----------------------------------------------------
        # Add every sheet/header
        # ----------------------------------------------------

        for sheet_name, headers in (
            self.sheet_headers.items()
        ):

            for header_info in headers:

                item_id = self.tree.insert(
                    "",
                    "end",
                    values=(
                        sheet_name,
                        self.column_letter(
                            header_info[
                                "column_index"
                            ]
                        ),
                        header_info[
                            "header"
                        ],
                        header_info[
                            "color"
                        ]
                    )
                )

                self.tree_metadata[
                    item_id
                ] = (
                    sheet_name,
                    header_info
                )

                self.update_tree_item_color(
                    item_id,
                    header_info[
                        "color"
                    ]
                )

    # ========================================================
    # COLUMN NUMBER TO EXCEL LETTER
    # ========================================================

    def column_letter(
        self,
        number
    ):

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

    def on_header_select(
        self,
        event=None
    ):

        selected = self.tree.selection()

        if not selected:

            self.selected_item = None

            self.selected_header_label.config(
                text="None"
            )

            return

        item_id = selected[0]

        self.selected_item = item_id

        if item_id not in self.tree_metadata:
            return

        sheet_name, header_info = (
            self.tree_metadata[
                item_id
            ]
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

    def apply_color(
        self,
        event=None
    ):

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

        # ----------------------------------------------------
        # Save color
        # ----------------------------------------------------

        header_info["color"] = color

        # ----------------------------------------------------
        # Update Treeview
        # ----------------------------------------------------

        values = list(
            self.tree.item(
                self.selected_item,
                "values"
            )
        )

        values[3] = color

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

    def update_tree_item_color(
        self,
        item_id,
        color
    ):

        tag_name = (
            f"color_{color}"
        )

        # ----------------------------------------------------
        # Create tag if needed
        # ----------------------------------------------------

        if not self.tree.tag_has(
            tag_name
        ):

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
            (
                "Are you sure you want to clear "
                "all header color mappings?"
            )
        )

        if not answer:
            return

        # ----------------------------------------------------
        # Reset every color
        # ----------------------------------------------------

        for headers in (
            self.sheet_headers.values()
        ):

            for header_info in headers:

                header_info[
                    "color"
                ] = "None"

        self.populate_tree()

        self.color_var.set(
            "None"
        )

        self.selected_header_label.config(
            text="None"
        )

        self.status_var.set(
            "All color mappings cleared."
        )

    # ========================================================
    # FIND COLOR CONFLICTS
    # ========================================================

    def get_color_conflicts(self):

        color_to_headers = OrderedDict()

        # ----------------------------------------------------
        # Collect header names for every color
        # ----------------------------------------------------

        for sheet_name, headers in (
            self.sheet_headers.items()
        ):

            for header_info in headers:

                color = header_info[
                    "color"
                ]

                if color == "None":
                    continue

                header = (
                    header_info[
                        "original_header"
                    ]
                )

                if color not in color_to_headers:

                    color_to_headers[
                        color
                    ] = []

                if (
                    header
                    not in color_to_headers[
                        color
                    ]
                ):

                    color_to_headers[
                        color
                    ].append(
                        header
                    )

        # ----------------------------------------------------
        # Find conflicts
        # ----------------------------------------------------

        conflicts = OrderedDict()

        for color, headers in (
            color_to_headers.items()
        ):

            if len(headers) > 1:

                conflicts[
                    color
                ] = headers

        return conflicts

    # ========================================================
    # VALIDATE MAPPING
    # ========================================================
    #
    # IMPORTANT:
    #
    # This function DOES NOT BLOCK the user.
    #
    # If same color has different headers:
    #     WARNING
    #
    # User can still consolidate.
    #
    # ========================================================

    def validate_mapping(self):

        if not self.sheet_headers:

            messagebox.showwarning(
                "No Workbook",
                (
                    "Please select an Excel "
                    "workbook first."
                )
            )

            return False

        conflicts = (
            self.get_color_conflicts()
        )

        # ----------------------------------------------------
        # Conflict found
        # ----------------------------------------------------

        if conflicts:

            conflict_text = "\n".join(
                f"• {color}: "
                f"{', '.join(headers)}"
                for color, headers
                in conflicts.items()
            )

            messagebox.showwarning(
                "Color Mapping Warning",
                (
                    "The same color is assigned "
                    "to different header names.\n\n"
                    f"{conflict_text}\n\n"
                    "This is only a warning.\n"
                    "You can still use FINAL "
                    "CONSOLIDATE."
                )
            )

            self.status_var.set(
                "Warning: same color used "
                "for different headers."
            )

            return True

        # ----------------------------------------------------
        # No conflict
        # ----------------------------------------------------

        messagebox.showinfo(
            "Mapping Valid",
            (
                "Header color mapping is valid.\n\n"
                "No conflicting color mapping "
                "was found."
            )
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

        # ----------------------------------------------------
        # Go through sheets in order
        # ----------------------------------------------------

        for sheet_name, headers in (
            self.sheet_headers.items()
        ):

            for header_info in headers:

                color = (
                    header_info[
                        "color"
                    ]
                )

                if color == "None":
                    continue

                header = (
                    header_info[
                        "original_header"
                    ]
                )

                if color not in color_groups:

                    color_groups[
                        color
                    ] = []

                if (
                    header
                    not in color_groups[
                        color
                    ]
                ):

                    color_groups[
                        color
                    ].append(
                        header
                    )

        return color_groups

    # ========================================================
    # BUILD OUTPUT HEADERS
    # ========================================================

    def build_output_headers(
        self,
        color_groups
    ):

        output_headers = []

        # ----------------------------------------------------
        # Colored groups
        #
        # First header found for a color becomes the
        # consolidated output header.
        # ----------------------------------------------------

        for color, headers in (
            color_groups.items()
        ):

            if headers:

                output_headers.append(
                    headers[0]
                )

        # ----------------------------------------------------
        # Uncolored headers
        # ----------------------------------------------------

        for sheet_name, headers in (
            self.sheet_headers.items()
        ):

            for header_info in headers:

                if (
                    header_info["color"]
                    != "None"
                ):
                    continue

                header = (
                    header_info[
                        "original_header"
                    ]
                )

                if (
                    header
                    not in output_headers
                ):

                    output_headers.append(
                        header
                    )

        return output_headers

    # ========================================================
    # GET OUTPUT HEADER FOR SOURCE HEADER
    # ========================================================

    def get_output_header(
        self,
        header_info,
        color_groups
    ):

        color = (
            header_info[
                "color"
            ]
        )

        original_header = (
            header_info[
                "original_header"
            ]
        )

        # ----------------------------------------------------
        # Colored header
        # ----------------------------------------------------

        if color != "None":

            if color in color_groups:

                # First header for that color
                # is the output header.
                return color_groups[
                    color
                ][0]

        # ----------------------------------------------------
        # Uncolored header
        # ----------------------------------------------------

        return original_header

    # ========================================================
    # FINAL CONSOLIDATION
    # ========================================================

    def final_consolidate(self):

        # ----------------------------------------------------
        # Check Excel file
        # ----------------------------------------------------

        if not self.file_path:

            messagebox.showwarning(
                "No Excel File",
                (
                    "Please select an Excel "
                    "workbook first."
                )
            )

            return

        # ----------------------------------------------------
        # Check sheets
        # ----------------------------------------------------

        if not self.sheet_headers:

            messagebox.showwarning(
                "No Data",
                "No worksheets were loaded."
            )

            return

        # ----------------------------------------------------
        # CHECK FOR COLOR CONFLICT
        # ----------------------------------------------------

        conflicts = (
            self.get_color_conflicts()
        )

        if conflicts:

            conflict_text = "\n".join(
                f"• {color}: "
                f"{', '.join(headers)}"
                for color, headers
                in conflicts.items()
            )

            # ------------------------------------------------
            # WARNING ONLY
            #
            # User can continue.
            # ------------------------------------------------

            answer = messagebox.askyesno(
                "Color Mapping Warning",
                (
                    "WARNING\n\n"
                    "The same color has been "
                    "assigned to different "
                    "header names.\n\n"
                    f"{conflict_text}\n\n"
                    "The program will treat headers "
                    "with the same color as the same "
                    "consolidated field.\n\n"
                    "Do you still want to continue "
                    "with FINAL CONSOLIDATE?"
                )
            )

            if not answer:

                self.status_var.set(
                    "Consolidation cancelled."
                )

                return

        # ----------------------------------------------------
        # FINAL CONFIRMATION
        # ----------------------------------------------------

        answer = messagebox.askyesno(
            "Final Consolidation",
            (
                "Ready to consolidate the workbook.\n\n"
                "The program will:\n\n"
                "• Preserve every source row\n"
                "• Append rows from all sheets\n"
                "• Align columns using header/color mapping\n"
                "• NOT match people or records\n"
                "• NOT deduplicate rows\n"
                "• NOT merge rows\n\n"
                "Do you want to continue?"
            )
        )

        if not answer:
            return

        # ----------------------------------------------------
        # OUTPUT FILE NAME
        # ----------------------------------------------------

        input_name = os.path.splitext(
            os.path.basename(
                self.file_path
            )
        )[0]

        default_name = (
            f"{input_name}_Consolidated.xlsx"
        )

        # ----------------------------------------------------
        # SELECT OUTPUT FILE
        # ----------------------------------------------------

        output_path = (
            filedialog.asksaveasfilename(
                title=(
                    "Save Consolidated "
                    "Excel File"
                ),
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[
                    (
                        "Excel Workbook",
                        "*.xlsx"
                    )
                ]
            )
        )

        if not output_path:
            return

        self.output_path = output_path

        # ----------------------------------------------------
        # START PROCESSING
        # ----------------------------------------------------

        try:

            self.status_var.set(
                "Consolidating data..."
            )

            self.root.update_idletasks()

            # ------------------------------------------------
            # Open Excel
            # ------------------------------------------------

            excel = pd.ExcelFile(
                self.file_path
            )

            # ------------------------------------------------
            # Store final rows
            # ------------------------------------------------

            all_output_rows = []

            # ------------------------------------------------
            # Color groups
            # ------------------------------------------------

            color_groups = (
                self.build_color_groups()
            )

            # ------------------------------------------------
            # Output headers
            # ------------------------------------------------

            output_headers = (
                self.build_output_headers(
                    color_groups
                )
            )

            # ------------------------------------------------
            # Total rows
            # ------------------------------------------------

            total_rows = 0

            # ------------------------------------------------
            # PROCESS EVERY SHEET
            # ------------------------------------------------

            for sheet_name in (
                excel.sheet_names
            ):

                self.status_var.set(
                    f"Processing sheet: "
                    f"{sheet_name}"
                )

                self.root.update_idletasks()

                # --------------------------------------------
                # Read sheet
                # --------------------------------------------

                df = pd.read_excel(
                    self.file_path,
                    sheet_name=sheet_name,
                    dtype=object
                )

                # --------------------------------------------
                # Empty sheet
                # --------------------------------------------

                if len(df.columns) == 0:
                    continue

                # --------------------------------------------
                # Header information for this sheet
                # --------------------------------------------

                sheet_header_info = (
                    self.sheet_headers.get(
                        sheet_name,
                        []
                    )
                )

                # --------------------------------------------
                # Column -> Output Header
                # --------------------------------------------

                column_output_map = {}

                for info in sheet_header_info:

                    column_index = (
                        info[
                            "column_index"
                        ]
                    )

                    if (
                        column_index
                        >= len(df.columns)
                    ):
                        continue

                    output_header = (
                        self.get_output_header(
                            info,
                            color_groups
                        )
                    )

                    column_output_map[
                        column_index
                    ] = output_header

                # --------------------------------------------
                # PROCESS EVERY ROW
                #
                # IMPORTANT:
                #
                # Every row remains a separate row.
                #
                # There is NO:
                #
                # - ID matching
                # - Name matching
                # - TIN matching
                # - duplicate removal
                # - row merging
                #
                # --------------------------------------------

                for _, source_row in (
                    df.iterrows()
                ):

                    output_row = {
                        header: ""
                        for header
                        in output_headers
                    }

                    has_data = False

                    # ----------------------------------------
                    # Copy values into correct output columns
                    # ----------------------------------------

                    for (
                        column_index,
                        output_header
                    ) in (
                        column_output_map.items()
                    ):

                        if (
                            column_index
                            >= len(source_row)
                        ):
                            continue

                        value = (
                            source_row.iloc[
                                column_index
                            ]
                        )

                        # ------------------------------------
                        # Keep non-empty values
                        # ------------------------------------

                        if pd.notna(value):

                            output_row[
                                output_header
                            ] = value

                            has_data = True

                    # ----------------------------------------
                    # Keep source row
                    # ----------------------------------------

                    if has_data:

                        all_output_rows.append(
                            output_row
                        )

                        total_rows += 1

            # ------------------------------------------------
            # CREATE FINAL DATAFRAME
            # ------------------------------------------------

            result_df = pd.DataFrame(
                all_output_rows,
                columns=output_headers
            )

            # ------------------------------------------------
            # SAVE OUTPUT
            # ------------------------------------------------

            # ------------------------------------------------
            # SAVE OUTPUT
            # Excel worksheet limit = 1,048,576 rows
            # Keep 1,048,575 rows for data because row 1 is header.
            # If data is larger, automatically create Part 1, Part 2, etc.
            # ------------------------------------------------

            MAX_EXCEL_ROWS = 1048576
            MAX_DATA_ROWS_PER_FILE = MAX_EXCEL_ROWS - 1

            if len(result_df) <= MAX_DATA_ROWS_PER_FILE:

                result_df.to_excel(
                    output_path,
                    index=False,
                    sheet_name="Consolidated"
                )

            else:

                base_name = os.path.splitext(output_path)[0]

                total_parts = (
                    (len(result_df) + MAX_DATA_ROWS_PER_FILE - 1)
                    // MAX_DATA_ROWS_PER_FILE
                )

                for part_number in range(1, total_parts + 1):

                    start_row = (
                        (part_number - 1)
                        * MAX_DATA_ROWS_PER_FILE
                    )

                    end_row = min(
                        start_row + MAX_DATA_ROWS_PER_FILE,
                        len(result_df)
                    )

                    part_df = result_df.iloc[
                        start_row:end_row
                    ]

                    part_output_path = (
                        f"{base_name}_Part_{part_number}.xlsx"
                    )

                    part_df.to_excel(
                        part_output_path,
                        index=False,
                        sheet_name="Consolidated"
                    )

            # ------------------------------------------------
            # SUCCESS
            # ------------------------------------------------

            self.status_var.set(
                (
                    f"Completed: "
                    f"{total_rows:,} rows consolidated."
                )
            )

            messagebox.showinfo(
                "Consolidation Complete",
                (
                    "Consolidation completed "
                    "successfully.\n\n"
                    f"Sheets processed: "
                    f"{len(excel.sheet_names)}\n"
                    f"Output columns: "
                    f"{len(output_headers)}\n"
                    f"Output rows: "
                    f"{total_rows:,}\n\n"
                    f"Saved to:\n"
                    f"{output_path}"
                )
            )

            # ------------------------------------------------
            # OPEN OUTPUT FOLDER
            # ------------------------------------------------

            open_folder = messagebox.askyesno(
                "Open Output Folder",
                (
                    "Do you want to open "
                    "the output folder?"
                )
            )

            if open_folder:

                folder = os.path.dirname(
                    output_path
                )

                try:

                    os.startfile(
                        folder
                    )

                except Exception:
                    pass

        # ----------------------------------------------------
        # PERMISSION ERROR
        # ----------------------------------------------------

        except PermissionError:

            messagebox.showerror(
                "Permission Error",
                (
                    "Unable to save the output file.\n\n"
                    "Please make sure the output Excel "
                    "file is not already open."
                )
            )

            self.status_var.set(
                "Unable to save output file."
            )

        # ----------------------------------------------------
        # GENERAL ERROR
        # ----------------------------------------------------

        except Exception as e:

            messagebox.showerror(
                "Consolidation Error",
                (
                    "An error occurred during "
                    "consolidation.\n\n"
                    f"{type(e).__name__}: {e}"
                )
            )

            self.status_var.set(
                "Consolidation failed."
            )


# ============================================================
# MAIN
# ============================================================

def main():

    root = tk.Tk()

    app = HeaderConsolidatorApp(
        root
    )

    root.mainloop()


# ============================================================
# START PROGRAM
# ============================================================

if __name__ == "__main__":

    main()
