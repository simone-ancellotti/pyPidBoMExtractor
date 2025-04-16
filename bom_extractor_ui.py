import tkinter as tk
from tkinter import ttk, filedialog, messagebox, PhotoImage
import logging
# Import the pyPidBoMExtractor package
from pyPidBoMExtractor.bom_generator import export_bom_to_excel, extract_bom_from_dxf,filterBOM_Ignore
from pyPidBoMExtractor.bom_generator import compare_bomsJSON,convert_bom_dxf_to_JSON,load_bom_from_excel_to_JSON
from pyPidBoMExtractor.bom_generator import header_mapping, header_mapping_reverse ,make_color_mapping,find_duplicates,update_XLS_add_missing_items_highlight
from pyPidBoMExtractor.importerdxf import import_BOMjson_into_DXF
from pyPidBoMExtractor.filterable_table import FilterableTable
import os
import json
import openpyxl

# pyinstaller --onefile --noconsole --strip --exclude-module=numpy bom_extractor_ui.py


# Configure logging to display in terminal
logging.basicConfig(level=logging.INFO)

# Main application class
class BOMExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BOM Extractor Application")
        self.geometry("610x440")  # Adjusted window size for better layout
        
        icon_image = PhotoImage(file="bom_valve_icon.png")
        self.iconphoto(True, icon_image)
        
        # Variables to hold file paths and options
        self.dwg_file = None
        self.template_BOM_xls_path = None
        self.revised_excel_file = None
        self.bom_dxf = None
        self.docDxf = None
        self.bom_revisedJSON = {}
        self.bom_dxf_JSON_like_xls= {}
        self.colour_mapping1 = None
        self.colour_mapping2 = None
        self.missing_in_dxf_color="FF0000"         # Default red
        self.missing_in_revised_color="CCCCCC"      # Default gray
        self.duplicate_color="800080"                # Default purple
        self.missing_in_revised = None
        self.missing_in_dxf = None
        self.workbook_excel = None
        self.flagExcelAldreadyCompared = False
        self.highlight_missing = tk.BooleanVar()  # Variable for highlight checkbox
        self.import_missing = tk.BooleanVar()  # Variable for import missing checkbox
        self.highlight_duplicate = tk.BooleanVar(value=True)  # Variable for highlight duplicate
        self.flagSaveNewExcellFile = tk.BooleanVar(value=True) 
        self.flagIgnoreWETEFE = tk.BooleanVar(value=False)
        
        #Create a Notebook (tabs container)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)
        
        # Create two tabs: one for controls, one for BOM table display.
        self.main_tab = ttk.Frame(self.notebook)
        self.table_dxf_tab = ttk.Frame(self.notebook)
        self.table_rev_tab = ttk.Frame(self.notebook)
        self.table_missing_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.main_tab, text="Main")
        self.notebook.add(self.table_dxf_tab, text="BOM Table dxf")
        self.notebook.add(self.table_rev_tab, text="BOM Table revised")
        #self.notebook.add(self.table_missing_tab, text="import DXF")
        
        
        # UI Setup
        self.setup_ui()
        self.create_menu() 
        #self.setup_table_tab()
        
        
        display_columns = list(header_mapping.values())
        self.table_dxf_items_filterable = FilterableTable(
            master=self.table_dxf_tab,
            data=self.bom_dxf_JSON_like_xls or {},   # Initially empty or loaded data.
            columns=display_columns,
            mapping=header_mapping_reverse,    # Map from display names to keys.
            filter_column_default="P&ID TAG",
            colour_mapping = None,
        )
        self.table_dxf_items_filterable.pack(fill="both", expand=True)
        
        display_columns = list(header_mapping.values())
        # In your setup method (e.g., setup_table_tab for the DXF tab):
        self.table_rev_tab_filterable = FilterableTable(
            master=self.table_rev_tab,
            data=self.bom_revisedJSON or {},   # Initially empty or loaded data.
            columns=display_columns,
            mapping=None,    # Map from display names to keys.
            filter_column_default="P&ID TAG",
            colour_mapping = None,
        )
        self.table_rev_tab_filterable.pack(fill="both", expand=True)
        

        
        self.bind_all("<ButtonRelease-1>", self.on_global_button_release)

    def on_global_button_release(self, event):
        """Global callback executed on every left mouse button release."""
        self.check_general_buttons()
        
    def create_menu(self):
        # Create the menu bar
        menubar = tk.Menu(self)
        
        # Create the File menu.
        file_menu = tk.Menu(menubar, tearoff=0)
        
        file_menu.add_command(label="Load settings", command=self.load_settings)
        file_menu.add_command(label="Save settings", command=self.save_settings)
        file_menu.add_separator()
        # Create a "Save" submenu.
        save_menu = tk.Menu(file_menu, tearoff=0)
        
        save_menu.add_separator()
        save_menu.add_command(label="Save revised xls BOM", command=self.exportNewExcellFile)
        save_menu.add_separator()
        save_menu.add_command(label="Save updated DXF", command=self.save_dxf_windows)
        
        # Add the Save cascade to the File menu.
        file_menu.add_cascade(label="Save", menu=save_menu)
        
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        
        # Add File menu to the menubar.
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Create a Help menu.
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.config(menu=menubar)
        
    def show_about(self):
        # Show an About dialog with version information
        about_text = "pyPidBoMExtractor Version 2.1\nDeveloped by Simone Ancellotti\n© 2025"
        messagebox.showinfo("About", about_text)

    def save_settings(self):
        """Save the current settings to a JSON file."""
        settings = {
            "dwg_file": self.dwg_file,
            "template_BOM_xls_path": self.template_BOM_xls_path,
            "revised_excel_file": self.revised_excel_file,
            #"bom_dxf": self.bom_dxf,  # If this is serializable (consider omitting if large)
            "highlight_missing": self.highlight_missing.get(),
            "import_missing": self.import_missing.get(),
            "highlight_duplicate": self.highlight_duplicate.get(),
            "flagSaveNewExcellFile": self.flagSaveNewExcellFile.get(),
            "flagIgnoreWETEFE": self.flagIgnoreWETEFE.get()
        }
        
        file = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            title="Save Settings As"
        )
        if file:
            try:
                with open(file, "w") as f:
                    json.dump(settings, f, indent=4)
                messagebox.showinfo("Settings Saved", "Settings have been successfully saved.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save settings: {e}")

    def load_settings(self):
        """Load settings from a JSON file and update the UI variables."""
        file = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            title="Load Settings"
        )
        if file:
            try:
                with open(file, "r") as f:
                    settings = json.load(f)
                
                # Update settings variables
                self.dwg_file = settings.get("dwg_file")
                self.template_BOM_xls_path = settings.get("template_BOM_xls_path")
                self.revised_excel_file = settings.get("revised_excel_file")
                #self.bom_dxf = settings.get("bom_dxf")
                self.highlight_missing.set(settings.get("highlight_missing", False))
                self.import_missing.set(settings.get("import_missing", False))
                self.highlight_duplicate.set(settings.get("highlight_duplicate", False))
                self.flagSaveNewExcellFile.set(settings.get("flagSaveNewExcellFile", True))
                self.flagIgnoreWETEFE.set(settings.get("flagIgnoreWETEFE", False))
                
                # Optionally, update your UI labels to reflect loaded file paths:
                if self.dwg_file:
                    self.dxf_label.config(text=os.path.basename(self.dwg_file))
                if self.template_BOM_xls_path:
                    self.template_label.config(text=os.path.basename(self.template_BOM_xls_path))
                if self.revised_excel_file:
                    self.revised_label.config(text=os.path.basename(self.revised_excel_file))
                
                self.check_ready_to_extract()
                #self.check_ready_to_exportBOM1()

                    
                messagebox.showinfo("Settings Loaded", "Settings have been successfully loaded.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load settings: {e}")
        
    def setup_ui(self):
        # Use a grid layout for better positioning
        

        # Upload DXF Button and Label (Top-left)
        self.dxf_button = tk.Button(self.main_tab, text="Upload DXF", command=self.upload_dxf)
        self.dxf_button.grid(row=0, column=0, padx=20, pady=10, sticky='w')

        self.dxf_label = tk.Label(self.main_tab, text="No DXF file uploaded")
        self.dxf_label.grid(row=1, column=0, padx=20, pady=5, sticky='w')

        # Upload Template Button and Label (Top-right)
        self.template_button = tk.Button(self.main_tab, text="Upload Template Excel", command=self.upload_template)
        self.template_button.grid(row=0, column=1, padx=20, pady=10, sticky='e')

        self.template_label = tk.Label(self.main_tab, text="No Template Excel uploaded")
        self.template_label.grid(row=1, column=1, padx=20, pady=5, sticky='e')
        
        # Checkbox to select whether to import missing DXF items into Excel (Center)
        self.import_checkbox = tk.Checkbutton(self.main_tab, text="Ignore WE TE FE", variable=self.flagIgnoreWETEFE)
        self.import_checkbox.grid(row=2, column=0, padx=20, pady=10, sticky='e')

        # # Extract BOM Button (Center)
        # self.extract_button = tk.Button(self, text="Extract BOM to Excel", state=tk.DISABLED, command=self.extract_bom)
        # self.extract_button.grid(row=2, column=0, columnspan=2, padx=20, pady=20)
        # Extract BOM Button
        self.extract_button = tk.Button(self.main_tab, text="Extract BOM from DXF", state=tk.DISABLED, command=self.extract_bom)
        self.extract_button.grid(row=3, column=0, columnspan=2, padx=20, pady=10,  sticky='w') 
        
        # Export BOM to Excel Button (initially disabled)
        self.export_button = tk.Button(self.main_tab, text="Export to Excel", state=tk.DISABLED, command=self.export_to_excel)
        self.export_button.grid(row=3, column=1, columnspan=2, padx=20, pady=10,  sticky='e')

        # Upload Revised Excel Button and Label (Center)
        self.revised_button = tk.Button(self.main_tab, text="Upload Revised BOM Excel", command=self.upload_revised_bom)
        self.revised_button.grid(row=4, column=0, padx=20, pady=10, sticky='w')

        self.revised_label = tk.Label(self.main_tab, text="No Revised BOM uploaded")
        self.revised_label.grid(row=4, column=1, padx=20, pady=10, sticky='e')

        # Checkbox to select whether to highlight missing components (Center) 
        self.highlight_checkbox = tk.Checkbutton(self.main_tab, text="Highlight in RED component \n in revised BOM but not in DXF",
                                                 variable=self.highlight_missing,
                                                 command=self.updateTableRevBOM)
        self.highlight_checkbox.grid(row=5, column=0, padx=20, pady=10, sticky='w')

        # Checkbox to select whether to import missing DXF items into Excel (Center)
        self.import_checkbox = tk.Checkbutton(self.main_tab, text="Import DXF items, which are \n missing in revised BOM,\n into new Excel. Highlight in GREY", 
                                              variable=self.import_missing,
                                              command=self.updateTableRevBOM)
        self.import_checkbox.grid(row=6, column=0, padx=20, pady=10, sticky='e')
        
        self.import_checkbox = tk.Checkbutton(self.main_tab, text="Highlight in PURPLE \n duplicated items in Excel", 
                                              variable=self.highlight_duplicate,
                                              command=self.updateTableRevBOM)
        self.import_checkbox.grid(row=5, column=1, padx=20, pady=10, sticky='e')
        
        # Checkbox to select whether save new file or modifiy exisiting file
        self.import_checkbox = tk.Checkbutton(self.main_tab, text="Save as new updated excell File", variable=self.flagSaveNewExcellFile)
        self.import_checkbox.grid(row=6, column=1, padx=20, pady=10, sticky='e')

        # Button to Compare BOM (Bottom-Center)
        self.compare_button = tk.Button(self.main_tab, text="Compare BOM vs DXF", state=tk.DISABLED, command=self.compare_bom)
        self.compare_button.grid(row=8, column=0, columnspan=1, padx=10, pady=20)
        
        
        self.import_dxf_button = tk.Button(self.main_tab, text="Import BOM into DXF", state=tk.DISABLED, command=self.import_BOM_into_DXF)
        self.import_dxf_button.grid(row=8, column=1, columnspan=1, padx=20, pady=20)
        
            
    # def upload_dxf(self):
    #     self.dwg_file = filedialog.askopenfilename(filetypes=[("DXF files", "*.dxf")])
    #     if self.dwg_file:
    #         self.dxf_label.config(text=os.path.basename(self.dwg_file))
    #         logging.info(f"Uploaded DXF file: {self.dwg_file}")
    #         self.extract_button.config(state=tk.NORMAL)
    
    # def setup_table_tab(self):
    #     # Create a frame to hold the table and its scrollbars
    #     table_frame = ttk.Frame(self.table_dxf_tab)
    #     table_frame.pack(fill="both", expand=True, padx=10, pady=10)

    #     # Use one row from sample_bom_data to determine columns.
         
    #     #sample_row = header_mapping[next(iter(header_mapping))]
    #     columns = list(header_mapping.values())

    #     # Create a Treeview to display the BOM data.
    #     self.tree1 = ttk.Treeview(table_frame, columns=columns, show="headings")
    #     self.tree1.grid(row=0, column=0, sticky="nsew")

    #     # Set up column headings.
    #     for col in columns:
    #         self.tree1.heading(col, text=col)
    #         self.tree1.column(col, width=100, anchor="center")
        

    #     # Add vertical and horizontal scrollbars.
    #     vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree1.yview)
    #     vsb.grid(row=0, column=1, sticky="ns")
    #     self.tree1.configure(yscrollcommand=vsb.set)

    #     hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree1.xview)
    #     hsb.grid(row=1, column=0, sticky="ew")
    #     self.tree1.configure(xscrollcommand=hsb.set)

    #     table_frame.rowconfigure(0, weight=1)
    #     table_frame.columnconfigure(0, weight=1)
    # def setup_table_tab(self):
    #     # --- Filtering Controls ---
    #     # Create a frame at the top of the DXF tab for filter controls.
    #     filter_frame = ttk.Frame(self.table_dxf_tab)
    #     filter_frame.pack(fill="x", padx=10, pady=5)
        
    #     ttk.Label(filter_frame, text="Filter Column:").pack(side="left", padx=5)
        
    #     # The options will be the friendly names (the values in header_mapping)
    #     options = list(header_mapping.values())
    #     self.dxf_filter_col_var = tk.StringVar(value=options[0])
    #     self.column_combobox = ttk.Combobox(filter_frame, textvariable=self.dxf_filter_col_var,
    #                                         values=options, state="readonly", width=15)
    #     self.column_combobox.pack(side="left", padx=5)
    #     self.column_combobox.bind("<<ComboboxSelected>>", self.update_dxf_filter)
        
    #     ttk.Label(filter_frame, text="Filter:").pack(side="left", padx=5)
    #     self.dxf_filter_text_var = tk.StringVar()
    #     filter_entry = ttk.Entry(filter_frame, textvariable=self.dxf_filter_text_var)
    #     filter_entry.pack(side="left", fill="x", expand=True, padx=5)
    #     filter_entry.bind("<KeyRelease>", self.update_dxf_filter)
        
    #     # --- Table Setup ---
    #     # Create a frame to hold the Treeview and its scrollbars.
    #     table_frame = ttk.Frame(self.table_dxf_tab)
    #     table_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
    #     # Use the header mapping for the display order. Here, we'll show the friendly names.
    #     self.dxf_display_columns = list(header_mapping.values())
    #     # And store the "real" column keys (order) for data retrieval.
    #     self.dxf_data_order = list(header_mapping.keys())
        
    #     self.tree1 = ttk.Treeview(table_frame, columns=self.dxf_display_columns, show="headings")
    #     self.tree1.grid(row=0, column=0, sticky="nsew")
        
    #     for col in self.dxf_display_columns:
    #         self.tree1.heading(col, text=col)
    #         self.tree1.column(col, width=80, anchor="center")
            
    #     # Add vertical scrollbar.
    #     vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree1.yview)
    #     vsb.grid(row=0, column=1, sticky="ns")
    #     self.tree1.configure(yscrollcommand=vsb.set)
        
    #     # Add horizontal scrollbar.
    #     hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree1.xview)
    #     hsb.grid(row=1, column=0, sticky="ew")
    #     self.tree1.configure(xscrollcommand=hsb.set)
        
    #     table_frame.rowconfigure(0, weight=1)
    #     table_frame.columnconfigure(0, weight=1)
    
    # def update_dxf_filter(self, event=None):
    #     """
    #     Called when the filter text or selected column changes.
    #     Filters self.bom_dxf and repopulates self.tree1.
    #     """
    #     # If no BOM data is loaded, simply return.
    #     if not self.bom_dxf:
    #         return
        
    #     # Invert header_mapping to map from friendly (display) name back to the actual key.
    #     inv_map = {v: k for k, v in header_mapping.items()}
    #     selected_display = self.dxf_filter_col_var.get()
    #     # Determine the actual dictionary key for filtering.
    #     filter_key = inv_map.get(selected_display, self.dxf_data_order[0])
        
    #     filter_text = self.dxf_filter_text_var.get().lower()
    #     filtered = {}
    #     for row_id, row in self.bom_dxf.items():
    #         cell_value = str(row.get(filter_key, "")).lower()
    #         if filter_text in cell_value:
    #             filtered[row_id] = row
    #     # Store the filtered data (you might want to keep the original in self.bom_dxf).
    #     self.filtered_bom_dxf = filtered
    #     self.updateTableDXFFiltered()
    
    # def updateTableDXFFiltered(self):
    #     """
    #     Clears and repopulates self.tree1 using self.filtered_bom_dxf.
    #     """
    #     # Clear current items from the Treeview.
    #     for item in self.tree1.get_children():
    #         self.tree1.delete(item)
        
    #     # Use the filtered data if available.
    #     data = self.filtered_bom_dxf if hasattr(self, 'filtered_bom_dxf') else self.bom_dxf
    #     if not data:
    #         return
        
    #     # Insert each row into the tree.
    #     for row in data.values():
    #         # Create a tuple of values in the order defined by self.dxf_data_order.
    #         values = tuple(row.get(key, "") for key in self.dxf_data_order)
    #         # Example: (optional) add a tag if you want to color certain rows.
    #         tags = ()
    #         # For instance, if you want to highlight rows whose "P&ID TAG" starts with "pH6":
    #         if str(row.get("P&ID TAG", "")).startswith("pH6"):
    #             tags = ("highlight",)
    #         self.tree1.insert("", "end", values=values, tags=tags)
        
    #     # Configure tag for row coloring.
    #     self.tree1.tag_configure("highlight", background="lightblue")

    def update_color_mapping_2nd_table(self):
        self.colour_mapping2={}
        if self.highlight_missing.get():
            color_sharp = "#"+self.missing_in_dxf_color
            return make_color_mapping(self.colour_mapping2 ,self.missing_in_dxf,color_sharp)
        else: 
            return {}
        #make_color_mapping(self.colour_mapping2 ,self.missing_in_dxf,"#CCCCCC")
        
    def update_color_mapping_1st_table(self):
        self.colour_mapping1={}
        if self.import_missing.get():
            color_sharp = "#"+self.missing_in_revised_color
            return make_color_mapping(self.colour_mapping1 ,self.missing_in_revised,color_sharp)
        else:
            return {}
        
    def update_color_mapping_duplicates_table(self,bom_dict,colour_mapping):
        if self.highlight_duplicate.get() and bom_dict and isinstance(colour_mapping,dict):
            duplicates=find_duplicates(bom_dict, 'P&ID TAG')
            duplicates_tags = [pid_tag for pid_tag in duplicates.keys()]
            color_sharp = "#"+self.duplicate_color
            return make_color_mapping(colour_mapping ,duplicates_tags,color_sharp)
        else:
            return {}
        
        
        
    def updateTableRevBOM(self):
        #colour_mapping ={'WT204':'#FF0000'}
        self.colour_mapping1= self.update_color_mapping_1st_table()
        self.colour_mapping2= self.update_color_mapping_2nd_table()
        colour_mapping11 = self.update_color_mapping_duplicates_table(self.bom_dxf,self.colour_mapping1)
        colour_mapping22 = self.update_color_mapping_duplicates_table(self.bom_revisedJSON,self.colour_mapping2)
        if colour_mapping11:
            self.colour_mapping1.update(colour_mapping11)
        if colour_mapping22:
            self.colour_mapping2.update(colour_mapping22)
            
        self.table_dxf_items_filterable.set_data(self.bom_dxf,self.colour_mapping1)
        self.table_rev_tab_filterable.set_data(self.bom_revisedJSON,self.colour_mapping2)
        
        
        
    def updateTableDXF(self):
        # Insert sample data into the tree1.
        # if self.bom_dxf:
        #     columns = list(header_mapping.keys())
        #     for row in self.bom_dxf.values():
        #          values = tuple(row.get(col, "") for col in columns)
        #          self.tree1.insert("", "end", values=values)    
        if self.bom_dxf:
             self.table_dxf_items_filterable.set_data(self.bom_dxf)
                 
    def upload_dxf(self):
        self.dwg_file = filedialog.askopenfilename(filetypes=[("DXF files", "*.dxf")])
        if self.dwg_file:
            self.dxf_label.config(text=os.path.basename(self.dwg_file))
            logging.info(f"Uploaded DXF file: {self.dwg_file}")
        self.check_ready_to_extract()
        
    def upload_template(self):
        self.template_BOM_xls_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.template_BOM_xls_path:
            self.template_label.config(text=os.path.basename(self.template_BOM_xls_path))
            logging.info(f"Uploaded Excel template: {self.template_BOM_xls_path}")
            if self.bom_dxf:
                #self.export_button.config(state=tk.NORMAL)
                self.check_ready_to_exportBOM1()

    def upload_revised_bom(self):
        # Open file dialog to select a revised BOM Excel file
        new_imported_xls_file_rev = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if new_imported_xls_file_rev != '':
            self.revised_excel_file = new_imported_xls_file_rev
            self.flagExcelAldreadyCompared = False
        if self.revised_excel_file:
            logging.info(f"Uploaded Revised BOM Excel file: {self.revised_excel_file}")
            self.bom_revisedJSON = load_bom_from_excel_to_JSON(self.revised_excel_file)
            self.revised_label.config(text=os.path.basename(self.revised_excel_file))
            #self.compare_button.config(state=tk.NORMAL)
            self.check_ready_to_export_revised_BOM2()
            self.updateTableRevBOM()

    def check_ready_to_extract(self):
        # Enable "Extract BOM" button if both DXF and Template Excel are uploaded
        """Enable the extract button if a valid DXF file is available."""
        if self.dwg_file and os.path.exists(self.dwg_file):
            self.extract_button.config(state=tk.NORMAL)
        else:
            self.extract_button.config(state=tk.DISABLED)

    def check_ready_to_exportBOM1(self):
        # Enable "Extract BOM" button if both DXF and Template Excel are uploaded
        """Enable the extract button if a valid DXF file is available."""
        if self.bom_dxf and self.dwg_file and os.path.exists(self.dwg_file) and self.template_BOM_xls_path and os.path.exists(self.template_BOM_xls_path):
            self.export_button.config(state=tk.NORMAL)
        else:
            self.export_button.config(state=tk.DISABLED)
            
    def check_ready_to_export_revised_BOM2(self):
        if self.revised_excel_file and os.path.exists(self.revised_excel_file):
            if self.bom_dxf and self.dwg_file and os.path.exists(self.dwg_file):
                self.compare_button.config(state=tk.NORMAL)
                self.import_dxf_button.config(state=tk.NORMAL)
                     
                     
    def check_general_buttons(self):
        self.check_ready_to_extract()
        self.check_ready_to_exportBOM1()
        self.check_ready_to_export_revised_BOM2()
        
    def extract_bom(self):
           if not self.dwg_file:
               messagebox.showerror("Error", "Please upload a DXF file first.")
               return
        
           logging.info("Extracting BOM from DXF...")
           self.bom_dxf,self.docDxf = extract_bom_from_dxf(self.dwg_file)
           self.bom_dxf_JSON_like_xls = convert_bom_dxf_to_JSON(self.bom_dxf)
           
           if self.flagIgnoreWETEFE.get():
               tagsvaluesToIgnore = ['WE', 'TE', 'FE']
               print('filtering '+str(tagsvaluesToIgnore))
               tagToIgnore='targetObjectType'
               bom_filtered = filterBOM_Ignore(self.bom_dxf,tagToIgnore,tagsvaluesToIgnore)
               self.bom_dxf = bom_filtered
           
           # Enable the Export button once BOM is extracted
           self.check_ready_to_exportBOM1()
           #self.updateTableDXF()
           self.updateTableRevBOM()
           #self.export_button.config(state=tk.NORMAL)

    def export_to_excel(self):
        if not self.bom_dxf:
            messagebox.showerror("Error", "Please extract the BOM first.")
            return
    
        if not self.template_BOM_xls_path:
            messagebox.showerror("Error", "Please upload an Excel template first.")
            return
    
        # Default file name
        excel_from_dxf_name = os.path.basename(self.dwg_file)
        default_filename = os.path.splitext(excel_from_dxf_name)[0] + "_bom.xlsx"
    
        # Open the "Save As" dialog
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_filename,
            title="Save BOM as"
        )
    
        if output_path:
            try:
                export_bom_to_excel(self.bom_dxf, self.template_BOM_xls_path, 
                                    output_path,
                                    self.highlight_duplicate)
                messagebox.showinfo("Success", f"BOM successfully exported to {output_path}")
            except Exception as e:
                logging.error(f"Failed to export BOM: {e}")
                messagebox.showerror("Error", f"Failed to export BOM: {e}")
    

        
    def compare_bom(self):
        try:
            logging.info("Comparing BOM with DXF...")
            if self.bom_dxf is None:
                logging.error("BOM has not been extracted yet.")
                messagebox.showerror("Error", "BOM must be extracted before comparison.")
                return
    
            # Load the revised Excel file as JSON
            self.bom_revisedJSON = load_bom_from_excel_to_JSON(self.revised_excel_file)
    
            # Convert BOM from DXF to JSON
            self.bom_dxf_JSON_like_xls = convert_bom_dxf_to_JSON(self.bom_dxf)
    

    
            # Perform BOM comparison
            # self.missing_in_revised, self.missing_in_dxf, self.workbook_excel = compare_bomsJSON(
            #     self.bom_dxf_JSON_like_xls,
            #     self.bom_revisedJSON,
            #     revised_excel_file=output_path,  # Save to the selected path
            #     highlight_duplicate=highlight_duplicate, 
            #     highlight_missing=highlight_missing,
            #     import_missingDXF2BOM=import_missingDXF2BOM
            #     )
            # Perform BOM comparison without modifying the xls file
            self.missing_in_revised, self.missing_in_dxf = compare_bomsJSON(
                self.bom_dxf_JSON_like_xls,
                self.bom_revisedJSON,
                # highlight_duplicate=highlight_duplicate, 
                # highlight_missing=highlight_missing,
                # import_missingDXF2BOM=import_missingDXF2BOM
                )
            
            self.flagExcelAldreadyCompared = True
                
            #self.update_color_mapping_2nd_table()
            #print(self.colour_mapping2)
            
            self.updateTableRevBOM()

            
            # print('missing_in_revised: '+str(missing_in_revised))
            # print('missing_in_dxf: '+str(missing_in_dxf))
            # Display comparison results
            messagebox.showinfo("Comparison Results", f"Missing in Revised: {self.missing_in_revised}\nMissing in DXF: {self.missing_in_dxf}")
            
            # if flagSaveNewExcellFile:
            #     rev_excel_file_name = os.path.basename(self.revised_excel_file)
            #     default_filename = os.path.splitext(rev_excel_file_name)[0] + "_updated.xlsx"
            #     output_path = filedialog.asksaveasfilename(
            #         defaultextension=".xlsx",
            #         filetypes=[("Excel files", "*.xlsx")],
            #         initialfile=default_filename,
            #         title="Save updated BOM as"
            #     )
            #     workbook_excel.save(output_path)  # Save changes to the Excel file
            #     if not output_path:
            #         messagebox.showerror("Error", "No file selected for saving the updated BOM.")
            #         return
            # else:
            #     workbook_excel.save(output_path)  # Save changes to the Excel file
                
           
        except Exception as e:
            logging.error(f"An error occurred during BOM comparison: {e}")
            messagebox.showerror("Error", f"Failed to compare BOM: {e}")

    def exportNewExcellFile(self):
        #if self.workbook_excel is None:
        if not(self.flagExcelAldreadyCompared):
            logging.error("BOM has not been compared yet.")
            messagebox.showerror("Error", "BOM must be compared before saving.")
            return
        # Check if the user wants to highlight missing components
        highlight_missing = self.highlight_missing.get()
        
        # Check if the user wants to import missing items from DXF to Excel
        import_missingDXF2BOM = self.import_missing.get()
        
        
        # Check if the user wants to highlight duplicate components
        highlight_duplicate = self.highlight_duplicate.get()
        
        
        # If the "Save As New" checkbox is ticked, prompt for a new file name
        output_path = self.revised_excel_file
        self.workbook_excel = openpyxl.load_workbook(output_path)

        print("Preparing for saving revised xls BOM...")
        update_XLS_add_missing_items_highlight( 
                    workbook_xls=self.workbook_excel,  
                    bom_dxf = self.bom_dxf_JSON_like_xls,
                    bom_revisedJSON = self.bom_revisedJSON,
                    missing_in_revised = self.missing_in_revised, 
                    missing_in_dxf = self.missing_in_dxf,
                    highlight_duplicate=highlight_duplicate, 
                    highlight_missing=highlight_missing,
                    import_missingDXF2BOM=import_missingDXF2BOM,
                    missing_in_dxf_color="FF0000",         
                    missing_in_revised_color="CCCCCC",      
                    duplicate_color="800080"                
                    )
        
        
        flagSaveNewExcellFile = self.flagSaveNewExcellFile.get()


        
        # Check if the directory exists; if not, create it
        output_dir = os.path.dirname(output_path)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            logging.info(f"Created directory: {output_dir}")
            
        if flagSaveNewExcellFile:
            rev_excel_file_name = os.path.basename(self.revised_excel_file)
            default_filename = os.path.splitext(rev_excel_file_name)[0] + "_updated.xlsx"
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=default_filename,
                title="Save updated BOM as"
            )
            
            if not output_path or output_path == "":
                messagebox.showerror("Error", "No file selected for saving the updated BOM.")
                return
            else: 
                self.workbook_excel.save(output_path)  # Save changes to the Excel file
                print("Revised xls BOM has been saved in a new xls file.")
        else:
            self.workbook_excel.save(output_path)  # Save changes to the Excel file
            self.workbook.close()
            print("Revised xls BOM has been saved in the same imported xls file.")
    
    def save_dxf_windows(self):
        """
        Open a 'Save As' dialog to store the updated DXF file.
        
        It uses the current DXF file name (self.dwg_file) to create a default file name,
        then uses self.docDxf.saveas to write the file to the user-selected path.
        """
        # Ensure a DXF file is loaded.
        if not self.dwg_file:
            messagebox.showerror("Error", "No DXF file loaded.")
            return
        
        # Generate a default file name based on the input DXF filename.
        dxf_file_name = os.path.basename(self.dwg_file)
        default_filename = os.path.splitext(dxf_file_name)[0] + "_updated.dxf"
    
        # Open a file dialog for saving the DXF file.
        output_path = filedialog.asksaveasfilename(
            defaultextension=".dxf",
            filetypes=[("DXF Files", "*.dxf")],
            initialfile=default_filename,
            title="Save updated DXF file as"
        )
        
        # Check if the user provided a file name.
        if not output_path:
            messagebox.showerror("Error", "No file selected for saving the updated DXF.")
            return
        
        try:
            # Save the updated DXF file using the chosen path.
            self.docDxf.saveas(output_path)
            messagebox.showinfo("Success", f"DXF saved successfully:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save DXF file: {e}")

    def import_BOM_into_DXF(self):
        # Load the revised Excel file as JSON
        bom_revisedJSON = load_bom_from_excel_to_JSON(self.revised_excel_file)
        rows_xls_no = import_BOMjson_into_DXF(bom_revisedJSON,self.bom_dxf)
        
        #self.save_dxf_windows()
        return None   
    
if __name__ == "__main__":
    app = BOMExtractorApp()
    app.mainloop()
