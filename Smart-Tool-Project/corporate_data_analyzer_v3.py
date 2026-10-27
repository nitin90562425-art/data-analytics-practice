# ```python
# """
# Corporate Data Analyzer
# -----------------------
# A professional desktop GUI application for analyzing CSV and Excel files.

# Features:
# - CSV and Excel file selection
# - Separate Read button
# - Dataset information
# - Automatic text/numeric column detection
# - Dynamic Group By / Aggregation / Value Column dropdowns
# - GroupBy report generation
# - Report preview inside GUI
# - CSV / Excel report export
# - Bar / Column / Line / Pie charts
# - Chart preview inside GUI
# - PNG chart export
# - Error handling
# - PyInstaller-friendly

# Author: Data Analytics Desktop Tool
# """

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg



def build_report(df, group, value, aggregation, order="Value descending"):
    """Pure reporting function: preserves input and rejects invalid numeric values."""
    if group not in df.columns:
        raise ValueError("Group column not found.")
    agg = "mean" if aggregation == "average" else aggregation
    if agg not in {"sum", "mean", "max", "min", "median", "count", "row count"}:
        raise ValueError("Unsupported aggregation.")
    group_values = df[group]
    blank_groups = group_values.isna() | group_values.astype(str).str.strip().eq("")
    notes = []
    if blank_groups.any():
        notes.append(f"Excluded {int(blank_groups.sum())} rows with blank groups.")
    work = df.loc[~blank_groups].copy()
    if agg == "row count":
        result = work.groupby(group, sort=False, observed=True).size().reset_index(name="Row count")
    else:
        if value not in df.columns:
            raise ValueError("Value column not found.")
        raw = work[value]
        missing = raw.isna() | raw.astype(str).str.strip().eq("")
        if agg == "count":
            # Count populated values, including text; not numeric-only values.
            metric = (~missing).astype(int)
            result = metric.groupby(work[group], sort=False, observed=True).sum().reset_index(name=f"Count of {value}")
        else:
            if pd.api.types.is_datetime64_any_dtype(raw):
                raise ValueError("Dates cannot be used as numeric values.")
            converted = pd.to_numeric(raw.where(~missing), errors="coerce")
            invalid = (~missing) & converted.isna()
            if invalid.any():
                examples = raw[invalid].astype(str).head(3).tolist()
                raise ValueError(f"{int(invalid.sum())} non-numeric values in {value}: {examples}. Clean them or use count/row count. No values were silently discarded.")
            import numpy as np
            if np.isinf(converted.dropna().astype(float)).any():
                raise ValueError("Infinite values are not allowed.")
            if missing.any():
                notes.append(f"{int(missing.sum())} missing values excluded from {agg}; all-missing groups remain blank.")
            grouped = converted.groupby(work[group], sort=False, observed=True)
            metric = grouped.sum(min_count=1) if agg == "sum" else grouped.agg(agg)
            result = metric.reset_index(name=f"{agg.capitalize()} of {value}")
    metric_name = result.columns[1]
    if order == "Value descending":
        result = result.sort_values(metric_name, ascending=False, kind="stable")
    elif order == "Group ascending":
        result = result.sort_values(group, kind="stable")
    elif order != "Original order":
        raise ValueError("Unknown sort order.")
    return result.reset_index(drop=True), notes

class DataAnalyzerApp:
    def __init__(self, root):
        self.root = root

        # ---------------------------------------------------------
        # WINDOW CONFIGURATION
        # ---------------------------------------------------------
        self.root.title("Corporate Data Analyzer")
        self.root.geometry("1250x800")
        self.root.minsize(1050, 700)

        # Application data
        self.df = None
        self.report_df = None
        self.file_path = ""
        self.file_folder = ""

        # Matplotlib objects
        self.figure = None
        self.canvas = None

        # ---------------------------------------------------------
        # STYLE
        # ---------------------------------------------------------
        self.setup_style()

        # ---------------------------------------------------------
        # BUILD GUI
        # ---------------------------------------------------------
        self.create_widgets()
        self.sort_mode = tk.StringVar(value="Value descending")
        ttk.Label(self.root, text="Report/chart order:").pack(side="left", padx=15)
        ttk.Combobox(self.root, textvariable=self.sort_mode, state="readonly",
                     values=["Value descending", "Group ascending", "Original order"], width=22).pack(side="left")
        self.read_warnings = []
        self.aggregation_combo.bind("<<ComboboxSelected>>", self.update_value_options)
        self.group_column_combo.bind("<<ComboboxSelected>>", self.update_value_options)
        self.update_value_options()

    # =============================================================
    # STYLE
    # =============================================================
    def setup_style(self):
        style = ttk.Style()

        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 18, "bold")
        )

        style.configure(
            "Section.TLabel",
            font=("Segoe UI", 11, "bold")
        )

        style.configure(
            "Info.TLabel",
            font=("Segoe UI", 10)
        )

        style.configure(
            "TButton",
            font=("Segoe UI", 10),
            padding=7
        )

        style.configure(
            "Accent.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=8
        )

        style.configure(
            "TCombobox",
            padding=5
        )

    # =============================================================
    # GUI
    # =============================================================
    def create_widgets(self):

        # ---------------------------------------------------------
        # MAIN CONTAINER
        # ---------------------------------------------------------
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill="both", expand=True)

        # ---------------------------------------------------------
        # TITLE
        # ---------------------------------------------------------
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 12))

        ttk.Label(
            title_frame,
            text="Corporate Data Analyzer",
            style="Title.TLabel"
        ).pack(side="left")

        ttk.Label(
            title_frame,
            text="CSV / Excel Analysis Tool",
            style="Info.TLabel"
        ).pack(side="right")

        # ---------------------------------------------------------
        # FILE SECTION
        # ---------------------------------------------------------
        file_frame = ttk.LabelFrame(
            main_frame,
            text="1. Select Data File",
            padding=10
        )
        file_frame.pack(fill="x", pady=5)

        self.file_entry = ttk.Entry(
            file_frame,
            width=90
        )
        self.file_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 8)
        )

        ttk.Button(
            file_frame,
            text="Browse",
            command=self.browse_file
        ).pack(side="left", padx=4)

        ttk.Button(
            file_frame,
            text="Read",
            style="Accent.TButton",
            command=self.read_file
        ).pack(side="left", padx=4)

        ttk.Button(
            file_frame,
            text="Reset",
            command=self.reset_all
        ).pack(side="left", padx=4)

        # ---------------------------------------------------------
        # DATA INFORMATION SECTION
        # ---------------------------------------------------------
        info_frame = ttk.LabelFrame(
            main_frame,
            text="2. Dataset Information",
            padding=10
        )
        info_frame.pack(fill="x", pady=5)

        self.rows_label = ttk.Label(
            info_frame,
            text="Rows: -",
            style="Info.TLabel"
        )
        self.rows_label.pack(side="left", padx=15)

        self.columns_label = ttk.Label(
            info_frame,
            text="Columns: -",
            style="Info.TLabel"
        )
        self.columns_label.pack(side="left", padx=15)

        self.column_names_label = ttk.Label(
            info_frame,
            text="Column Headings: -",
            style="Info.TLabel",
            wraplength=900
        )
        self.column_names_label.pack(
            side="left",
            padx=15,
            fill="x",
            expand=True
        )

        # ---------------------------------------------------------
        # REPORT BUILDER
        # ---------------------------------------------------------
        builder_frame = ttk.LabelFrame(
            main_frame,
            text="3. Report Builder",
            padding=10
        )
        builder_frame.pack(fill="x", pady=5)

        # Group By
        ttk.Label(
            builder_frame,
            text="Group By Column:"
        ).grid(
            row=0,
            column=0,
            padx=5,
            pady=5,
            sticky="w"
        )

        self.group_column_combo = ttk.Combobox(
            builder_frame,
            state="readonly",
            width=25
        )
        self.group_column_combo.grid(
            row=0,
            column=1,
            padx=5,
            pady=5
        )

        # Aggregation
        ttk.Label(
            builder_frame,
            text="Aggregation:"
        ).grid(
            row=0,
            column=2,
            padx=5,
            pady=5,
            sticky="w"
        )

        self.aggregation_combo = ttk.Combobox(
            builder_frame,
            state="readonly",
            width=20,
            values=[
                "sum",
                "mean",
                "average",
                "max",
                "min",
                "count",
                "median",
                "row count"
            ]
        )
        self.aggregation_combo.grid(
            row=0,
            column=3,
            padx=5,
            pady=5
        )

        # Value Column
        ttk.Label(
            builder_frame,
            text="Value Column:"
        ).grid(
            row=0,
            column=4,
            padx=5,
            pady=5,
            sticky="w"
        )

        self.value_column_combo = ttk.Combobox(
            builder_frame,
            state="readonly",
            width=25
        )
        self.value_column_combo.grid(
            row=0,
            column=5,
            padx=5,
            pady=5
        )

        # Preview Report
        ttk.Button(
            builder_frame,
            text="Preview Report",
            style="Accent.TButton",
            command=self.preview_report
        ).grid(
            row=0,
            column=6,
            padx=10,
            pady=5
        )

        # ---------------------------------------------------------
        # REPORT EXPORT
        # ---------------------------------------------------------
        export_frame = ttk.Frame(builder_frame)
        export_frame.grid(
            row=1,
            column=0,
            columnspan=7,
            sticky="w",
            pady=(8, 0)
        )

        ttk.Label(
            export_frame,
            text="Export Format:"
        ).pack(side="left", padx=5)

        self.export_format_combo = ttk.Combobox(
            export_frame,
            state="readonly",
            width=15,
            values=[
                "Excel (.xlsx)",
                "CSV (.csv)"
            ]
        )
        self.export_format_combo.pack(
            side="left",
            padx=5
        )

        ttk.Button(
            export_frame,
            text="Export Report",
            command=self.export_report
        ).pack(
            side="left",
            padx=5
        )

        # ---------------------------------------------------------
        # MAIN CONTENT AREA
        # ---------------------------------------------------------
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(
            fill="both",
            expand=True,
            pady=5
        )

        # ---------------------------------------------------------
        # LEFT SIDE - REPORT TABLE
        # ---------------------------------------------------------
        report_frame = ttk.LabelFrame(
            content_frame,
            text="Report Preview",
            padding=8
        )
        report_frame.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(0, 5)
        )

        table_container = ttk.Frame(report_frame)
        table_container.pack(
            fill="both",
            expand=True
        )

        self.report_tree = ttk.Treeview(
            table_container,
            show="headings"
        )

        vertical_scroll = ttk.Scrollbar(
            table_container,
            orient="vertical",
            command=self.report_tree.yview
        )

        horizontal_scroll = ttk.Scrollbar(
            table_container,
            orient="horizontal",
            command=self.report_tree.xview
        )

        self.report_tree.configure(
            yscrollcommand=vertical_scroll.set,
            xscrollcommand=horizontal_scroll.set
        )

        self.report_tree.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        vertical_scroll.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        horizontal_scroll.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        table_container.rowconfigure(0, weight=1)
        table_container.columnconfigure(0, weight=1)

        # ---------------------------------------------------------
        # RIGHT SIDE - CHART
        # ---------------------------------------------------------
        chart_frame = ttk.LabelFrame(
            content_frame,
            text="Chart Preview",
            padding=8
        )
        chart_frame.pack(
            side="right",
            fill="both",
            expand=True,
            padx=(5, 0)
        )

        chart_controls = ttk.Frame(chart_frame)
        chart_controls.pack(
            fill="x",
            pady=(0, 5)
        )

        ttk.Label(
            chart_controls,
            text="Chart Type:"
        ).pack(side="left", padx=5)

        self.chart_type_combo = ttk.Combobox(
            chart_controls,
            state="readonly",
            width=18,
            values=[
                "Bar chart",
                "Column chart",
                "Line chart",
                "Pie chart"
            ]
        )
        self.chart_type_combo.pack(
            side="left",
            padx=5
        )

        ttk.Button(
            chart_controls,
            text="Preview Chart",
            command=self.preview_chart
        ).pack(
            side="left",
            padx=5
        )

        ttk.Button(
            chart_controls,
            text="Export Chart",
            command=self.export_chart
        ).pack(
            side="left",
            padx=5
        )

        # Chart area
        self.chart_container = ttk.Frame(chart_frame)
        self.chart_container.pack(
            fill="both",
            expand=True
        )

        # ---------------------------------------------------------
        # STATUS BAR
        # ---------------------------------------------------------
        self.status_var = tk.StringVar(
            value="Ready. Select a CSV or Excel file."
        )

        status_bar = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            relief="sunken",
            anchor="w",
            padding=5
        )
        status_bar.pack(
            fill="x",
            pady=(5, 0)
        )

    # =============================================================
    # FILE BROWSER
    # =============================================================
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV or Excel File",
            filetypes=[
                ("Data Files", "*.csv *.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx *.xls"),
                ("All Files", "*.*")
            ]
        )

        if not file_path:
            return

        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, file_path)

        # Reset old analysis when a new file is selected
        self.clear_analysis()

        self.status_var.set(
            "File selected. Click Read to load the dataset."
        )

    # =============================================================
    # READ FILE
    # =============================================================
    def read_file(self):
        file_path = self.file_entry.get().strip()

        if not file_path:
            messagebox.showerror(
                "File Required",
                "Please select a CSV or Excel file first."
            )
            return

        if not os.path.exists(file_path):
            messagebox.showerror(
                "File Not Found",
                "The selected file does not exist."
            )
            return

        self.clear_analysis()
        try:
            extension = os.path.splitext(file_path)[1].lower()

            # -----------------------------------------------------
            # READ CSV
            # -----------------------------------------------------
            if extension == ".csv":
                df = pd.read_csv(file_path, dtype=str, keep_default_na=False)

            # -----------------------------------------------------
            # READ EXCEL
            # -----------------------------------------------------
            elif extension in [".xlsx", ".xls"]:
                df = pd.read_excel(file_path, dtype=object, keep_default_na=False)

            else:
                messagebox.showerror(
                    "Unsupported File",
                    "Please select a CSV or Excel file."
                )
                return

            # -----------------------------------------------------
            # CHECK EMPTY DATASET
            # -----------------------------------------------------
            if df.empty:
                messagebox.showerror(
                    "Empty Dataset",
                    "The selected file does not contain any data."
                )
                return

            self.clear_analysis()
            self.df = df
            self.file_path = file_path
            self.file_folder = os.path.dirname(file_path)

            # -----------------------------------------------------
            # DETECT COLUMN TYPES
            # -----------------------------------------------------
            self.detect_columns()

            # -----------------------------------------------------
            # DISPLAY DATASET INFORMATION
            # -----------------------------------------------------
            self.rows_label.config(
                text=f"Rows: {len(self.df):,}"
            )

            self.columns_label.config(
                text=f"Columns: {len(self.df.columns):,}"
            )

            column_names = ", ".join(
                str(column)
                for column in self.df.columns
            )

            self.column_names_label.config(
                text=f"Column Headings: {column_names}"
            )

            self.status_var.set(
                f"Dataset loaded successfully: "
                f"{len(self.df):,} rows × {len(self.df.columns):,} columns"
            )

            messagebox.showinfo(
                "File Loaded",
                f"Dataset loaded successfully.\n\n"
                f"Rows: {len(self.df):,}\n"
                f"Columns: {len(self.df.columns):,}"
            )

        except Exception as error:
            messagebox.showerror(
                "Read Error",
                f"Unable to read the file.\n\n"
                f"Error: {error}"
            )

            self.status_var.set("Failed to read file.")

    # =============================================================
    # DETECT COLUMN TYPES
    # =============================================================
    def detect_columns(self):
        # Preserve loaded values. Type inference must not silently mutate data.
        self.text_columns = [c for c in self.df.columns
                             if not pd.api.types.is_numeric_dtype(self.df[c])]
        self.numeric_columns = [c for c in self.df.columns
                                if pd.api.types.is_numeric_dtype(self.df[c])]
        # Numeric IDs and dates can be valid grouping columns too.
        columns = list(self.df.columns)
        self.group_column_combo["values"] = columns
        self.value_column_combo["values"] = columns
        self.group_column_combo.set("")
        self.value_column_combo.set("")
        self.update_value_options()

    # =============================================================
    # PREVIEW REPORT
    # =============================================================
    def update_value_options(self, event=None):
        agg = self.aggregation_combo.get()
        if agg == "row count" or self.df is None:
            self.value_column_combo.set("")
            self.value_column_combo["values"] = []
            self.value_column_combo.configure(state="disabled")
            return
        self.value_column_combo.configure(state="readonly")
        group = self.group_column_combo.get()
        eligible = []
        for col in self.df.columns:
            raw = self.df[col]
            if agg == "count":
                eligible.append(col)
                continue
            if col == group or pd.api.types.is_datetime64_any_dtype(raw):
                continue
            missing = raw.isna() | raw.astype(str).str.strip().eq("")
            populated = raw[~missing]
            if not populated.empty:
                converted = pd.to_numeric(populated, errors="coerce")
                if converted.notna().all():
                    eligible.append(col)
        old = self.value_column_combo.get()
        self.value_column_combo["values"] = eligible
        if old not in eligible:
            self.value_column_combo.set("")

    def preview_report(self):
        self.clear_report()
        self.clear_chart()
        if self.df is None:
            messagebox.showerror("Read Required", "Please load a dataset first.")
            return
        group = self.group_column_combo.get()
        value = self.value_column_combo.get()
        agg = self.aggregation_combo.get()
        if not group or not agg or (agg != "row count" and not value):
            messagebox.showerror("Selection Required", "Select a group, aggregation and value column (except row count).")
            return
        try:
            order = self.sort_mode.get()
            result, notes = build_report(self.df, group, value, agg, order)
            if result.empty:
                raise ValueError("No groups remain after excluding blank group values.")
            self.report_df = result
            self.display_report(result)
            self.status_var.set(f"Report: {len(result):,} groups. " + " ".join(notes))
            if notes:
                messagebox.showwarning("Data quality notes", "\n".join(notes))
        except Exception as error:
            messagebox.showerror("Report Error", str(error))

    # =============================================================
    # DISPLAY REPORT IN TREEVIEW
    # =============================================================
    def display_report(self, dataframe):

        # ---------------------------------------------------------
        # REMOVE OLD TABLE DATA
        # ---------------------------------------------------------
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)

        # ---------------------------------------------------------
        # REMOVE OLD COLUMNS
        # ---------------------------------------------------------
        self.report_tree["columns"] = []

        # ---------------------------------------------------------
        # CREATE COLUMNS
        # ---------------------------------------------------------
        columns = list(dataframe.columns)

        self.report_tree["columns"] = columns

        for column in columns:

            self.report_tree.heading(
                column,
                text=str(column)
            )

            self.report_tree.column(
                column,
                width=180,
                minwidth=100,
                anchor="center"
            )

        # ---------------------------------------------------------
        # INSERT ROWS
        # ---------------------------------------------------------
        for _, row in dataframe.iterrows():

            values = []

            for value in row:

                if pd.isna(value):
                    values.append("")
                elif isinstance(value, float):
                    values.append(f"{value:,.2f}")
                else:
                    values.append(str(value))

            self.report_tree.insert(
                "",
                "end",
                values=values
            )

    # =============================================================
    # EXPORT REPORT
    # =============================================================
    def export_report(self):

        if self.df is None:
            messagebox.showerror(
                "Read Required",
                "Please read a file before exporting."
            )
            return

        if self.report_df is None:
            messagebox.showerror(
                "Report Required",
                "Please click Preview Report first."
            )
            return

        export_format = self.export_format_combo.get().strip()

        if not export_format:
            messagebox.showerror(
                "Format Required",
                "Please select an export format."
            )
            return

        try:

            # -----------------------------------------------------
            # INPUT FILE NAME
            # -----------------------------------------------------
            input_name = os.path.splitext(
                os.path.basename(self.file_path)
            )[0]

            # -----------------------------------------------------
            # EXPORT EXCEL
            # -----------------------------------------------------
            if export_format == "Excel (.xlsx)":

                output_path = filedialog.asksaveasfilename(
                    title="Save report", initialdir=self.file_folder,
                    initialfile=f"{input_name}_Report.xlsx", defaultextension=".xlsx",
                    filetypes=[("Excel", "*.xlsx")], confirmoverwrite=True)
                if not output_path:
                    return

                self.report_df.to_excel(
                    output_path,
                    index=False
                )

            # -----------------------------------------------------
            # EXPORT CSV
            # -----------------------------------------------------
            elif export_format == "CSV (.csv)":

                output_path = filedialog.asksaveasfilename(
                    title="Save report", initialdir=self.file_folder,
                    initialfile=f"{input_name}_Report.csv", defaultextension=".csv",
                    filetypes=[("CSV", "*.csv")], confirmoverwrite=True)
                if not output_path:
                    return

                self.report_df.to_csv(
                    output_path,
                    index=False
                )

            else:
                messagebox.showerror(
                    "Invalid Format",
                    "Please select a valid export format."
                )
                return

            self.status_var.set(
                f"Report exported: {output_path}"
            )

            messagebox.showinfo(
                "Export Successful",
                f"Report exported successfully.\n\n"
                f"Location:\n{output_path}"
            )

        except PermissionError:
            messagebox.showerror(
                "Permission Error",
                "Unable to save the report.\n\n"
                "The file may already be open in Excel. "
                "Please close it and try again."
            )

        except Exception as error:

            messagebox.showerror(
                "Export Error",
                f"Unable to export the report.\n\n"
                f"Error: {error}"
            )

    # =============================================================
    # PREVIEW CHART
    # =============================================================
    def preview_chart(self):

        if self.df is None:
            messagebox.showerror(
                "Read Required",
                "Please read a file first."
            )
            return

        if self.report_df is None:
            messagebox.showerror(
                "Report Required",
                "Please generate a report first."
            )
            return

        chart_type = self.chart_type_combo.get().strip()

        if not chart_type:
            messagebox.showerror(
                "Selection Required",
                "Please select a chart type."
            )
            return

        try:

            # -----------------------------------------------------
            # CLEAR PREVIOUS CHART
            # -----------------------------------------------------
            self.clear_chart()

            group_column = self.report_df.columns[0]
            value_column = self.report_df.columns[1]

            # -----------------------------------------------------
            # LIMIT PIE CHART CATEGORIES
            # -----------------------------------------------------
            chart_df = self.report_df.copy()

            if chart_type == "Pie chart":
                values = pd.to_numeric(chart_df[value_column], errors="raise")
                if values.isna().any() or (values < 0).any() or values.sum() <= 0:
                    raise ValueError("Pie chart needs non-missing, non-negative values with a positive total.")
                if self.aggregation_combo.get() not in ["sum", "count", "row count"]:
                    raise ValueError("Pie charts are restricted to additive metrics: sum, count or row count.")
                chart_df = chart_df.sort_values(value_column, ascending=False)
                if len(chart_df) > 10:
                    remainder = chart_df.iloc[10:][value_column].sum()
                    chart_df = pd.concat([chart_df.head(10), pd.DataFrame({
                        group_column: ["Other categories (combined)"],
                        value_column: [remainder]})], ignore_index=True)

            # Keep all report categories; size chart according to category count.
            category_count = len(chart_df)
            labels = chart_df[group_column].astype(str).tolist()
            positions = list(range(category_count))
            # Wrap long labels instead of letting them collide.
            import textwrap
            labels = ["\n".join(textwrap.wrap(label, 20)) for label in labels]
            self.root.update_idletasks()
            panel_width = max(300, self.chart_container.winfo_width())
            panel_height = max(240, self.chart_container.winfo_height())
            chart_size = (panel_width / 100, panel_height / 100)
            # -----------------------------------------------------
            # CREATE FIGURE
            # -----------------------------------------------------
            self.figure, ax = plt.subplots(
                figsize=chart_size, dpi=100, layout="constrained"
            )

            # -----------------------------------------------------
            # BAR CHART - HORIZONTAL
            # -----------------------------------------------------
            if chart_type == "Bar chart":

                ax.barh(
                    positions,
                    chart_df[value_column], height=0.62
                )
                ax.set_yticks(positions, labels=labels)
                ax.margins(y=0.04)
                ax.set_xlabel(value_column)
                ax.set_ylabel(group_column)

                ax.invert_yaxis()

            # -----------------------------------------------------
            # COLUMN CHART - VERTICAL
            # -----------------------------------------------------
            elif chart_type == "Column chart":

                ax.bar(
                    positions,
                    chart_df[value_column], width=0.62
                )
                ax.set_xticks(positions, labels=labels)
                ax.margins(x=0.04)
                ax.set_xlabel(group_column)
                ax.set_ylabel(value_column)

                plt.setp(
                    ax.get_xticklabels(),
                    rotation=45,
                    ha="right"
                )

            # -----------------------------------------------------
            # LINE CHART
            # -----------------------------------------------------
            elif chart_type == "Line chart":
                if self.sort_mode.get() == "Value descending":
                    raise ValueError("For line charts choose Group ascending (chronological dates) or Original order. Value-ranked lines do not show a time trend.")
                ax.plot(
                    positions,
                    chart_df[value_column],
                    marker="o"
                )
                ax.set_xticks(positions, labels=labels)

                ax.set_xlabel(group_column)
                ax.set_ylabel(value_column)

                plt.setp(
                    ax.get_xticklabels(),
                    rotation=45,
                    ha="right"
                )

            # -----------------------------------------------------
            # PIE CHART
            # -----------------------------------------------------
            elif chart_type == "Pie chart":

                ax.pie(
                    chart_df[value_column],
                    labels=chart_df[group_column].astype(str),
                    autopct="%1.1f%%",
                    startangle=90
                )

                ax.set_title(
                    f"{value_column} by {group_column}"
                )

            else:
                raise ValueError(
                    "Invalid chart type."
                )

            # -----------------------------------------------------
            # TITLE
            # -----------------------------------------------------
            if chart_type != "Pie chart":

                ax.set_title(
                    f"{value_column} by {group_column}"
                )

            # -----------------------------------------------------
            # GRID FOR NON-PIE CHARTS
            # -----------------------------------------------------
            if chart_type != "Pie chart":

                ax.grid(
                    axis="x" if chart_type == "Bar chart" else "y",
                    alpha=0.3
                )

            self.chart_axis = ax
            self.chart_labels = labels
            self.chart_positions = positions
            self.rendered_chart_type = chart_type
            self.adjust_chart_labels(panel_width, panel_height)

            # -----------------------------------------------------
            # EMBED MATPLOTLIB INTO TKINTER
            # -----------------------------------------------------
            self.canvas = FigureCanvasTkAgg(self.figure, master=self.chart_container)
            self.canvas.get_tk_widget().bind("<Configure>", self.on_chart_resize, add="+")

            self.canvas.draw()
            self.rendered_chart_type = chart_type

            self.canvas.get_tk_widget().pack(fill="both", expand=True)

            self.status_var.set(
                f"{chart_type}: fitted to panel; {category_count} plotted categories. "
                + ("Some labels are hidden to avoid overlap; all data points remain." if getattr(self, "chart_sparse_labels", False) and chart_type != "Pie chart" else "")
            )

        except Exception as error:
            self.clear_chart()
            messagebox.showerror(
                "Chart Error",
                f"Unable to generate chart.\n\n"
                f"Error: {error}"
            )

    # =============================================================
    # EXPORT CHART
    # =============================================================
    def adjust_chart_labels(self, width, height):
        ax = self.chart_axis
        chart_type = self.rendered_chart_type
        if chart_type == "Pie chart":
            for text in ax.texts:
                text.set_fontsize(8 if width < 600 else 9)
            return
        import math
        labels = self.chart_labels
        n = len(labels)
        # Preserve every bar/data point. Only thin tick labels when space is limited.
        capacity = max(2, int((height - 90) / 30)) if chart_type == "Bar chart" else max(2, int((width - 110) / 80))
        step = max(1, math.ceil(n / capacity))
        indices = list(range(0, n, step))
        if n and (not indices or indices[-1] != n - 1):
            indices.append(n - 1)
        import textwrap
        fitted = []
        for i in indices:
            text = labels[i].replace("\n", " ")
            parts = textwrap.wrap(text, 16 if chart_type == "Bar chart" else 12)
            fitted.append("\n".join(parts[:2]) + ("…" if len(parts) > 2 else ""))
        if chart_type == "Bar chart":
            ax.set_yticks(indices, labels=fitted)
        else:
            ax.set_xticks(indices, labels=fitted)
            plt.setp(ax.get_xticklabels(), rotation=35 if n > 4 else 0,
                     ha="right" if n > 4 else "center")
        ax.tick_params(axis="both", labelsize=8 if width < 600 else 9)
        ax.title.set_fontsize(10)
        ax.xaxis.label.set_fontsize(9)
        ax.yaxis.label.set_fontsize(9)
        self.chart_sparse_labels = step > 1

    def on_chart_resize(self, event):
        if self.figure is None or self.canvas is None:
            return
        # TkAgg's own Configure handler resizes the figure to this widget.
        self.adjust_chart_labels(event.width, event.height)
        self.canvas.draw_idle()

    def export_chart(self):

        if self.df is None:
            messagebox.showerror(
                "Read Required",
                "Please read a file first."
            )
            return

        if self.report_df is None:
            messagebox.showerror(
                "Report Required",
                "Please generate a report first."
            )
            return

        if self.figure is None:
            messagebox.showerror(
                "Chart Required",
                "Please click Preview Chart first."
            )
            return

        try:

            input_name = os.path.splitext(
                os.path.basename(self.file_path)
            )[0]

            chart_type = getattr(self, "rendered_chart_type", "Chart")

            safe_chart_type = (
                chart_type
                .replace(" ", "_")
                .replace("/", "_")
            )

            output_path = filedialog.asksaveasfilename(
                title="Save chart", initialdir=self.file_folder,
                initialfile=f"{input_name}_{safe_chart_type}.png", defaultextension=".png",
                filetypes=[("PNG", "*.png")], confirmoverwrite=True)
            if not output_path:
                return

            self.figure.savefig(
                output_path,
                dpi=300,
                bbox_inches="tight"
            )

            self.status_var.set(
                f"Chart exported: {output_path}"
            )

            messagebox.showinfo(
                "Export Successful",
                f"Chart exported successfully.\n\n"
                f"Location:\n{output_path}"
            )

        except PermissionError:

            messagebox.showerror(
                "Permission Error",
                "Unable to save the chart. "
                "Please check the folder permissions."
            )

        except Exception as error:

            messagebox.showerror(
                "Export Error",
                f"Unable to export the chart.\n\n"
                f"Error: {error}"
            )

    # =============================================================
    # CLEAR REPORT TABLE
    # =============================================================
    def clear_report(self):

        for item in self.report_tree.get_children():
            self.report_tree.delete(item)

        self.report_tree["columns"] = []

        self.report_df = None

    # =============================================================
    # CLEAR CHART
    # =============================================================
    def clear_chart(self):
        for widget in self.chart_container.winfo_children():
            widget.destroy()
        if self.canvas is not None:
            self.canvas = None

        if self.canvas is not None:

            self.canvas.get_tk_widget().destroy()
            self.canvas = None

        if self.figure is not None:

            plt.close(self.figure)
            self.figure = None

    # =============================================================
    # CLEAR ANALYSIS
    # =============================================================
    def clear_analysis(self):

        self.clear_report()
        self.clear_chart()

        self.rows_label.config(
            text="Rows: -"
        )

        self.columns_label.config(
            text="Columns: -"
        )

        self.column_names_label.config(
            text="Column Headings: -"
        )

        self.group_column_combo.set("")
        self.value_column_combo.set("")
        self.aggregation_combo.set("")
        self.export_format_combo.set("")
        self.chart_type_combo.set("")

        self.group_column_combo["values"] = []
        self.value_column_combo["values"] = []

        self.df = None
        self.file_path = ""
        self.file_folder = ""

    # =============================================================
    # RESET EVERYTHING
    # =============================================================
    def reset_all(self):

        self.file_entry.delete(
            0,
            tk.END
        )

        self.clear_analysis()

        self.status_var.set(
            "Ready. Select a CSV or Excel file."
        )


# =================================================================
# APPLICATION ENTRY POINT
# =================================================================
def main():

    root = tk.Tk()

    app = DataAnalyzerApp(root)

    root.mainloop()


if __name__ == "__main__":
    main()