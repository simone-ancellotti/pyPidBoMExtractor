import tkinter as tk
from tkinter import filedialog, messagebox
import logging
from pyPidBoMExtractor.bom_generator import load_bom_from_excel, extract_bom_from_dxf, export_bom_to_excel, convert_bom_dxf_to_dataframe, compare_boms
import os

# Configure logging to display in terminal
logging.basicConfig(level=logging.INFO)

# Main application class
class BOMExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BOM Extractor Application")
        self.geometry("600x400")  # Adjusted window size for better layout

        # Variables to hold file paths and options
        self.dwg_file = None
        self.template_BOM_xls_path = None
        self.revised_excel_file = None
        self.bom_dxf = None
        self.highlight_missing = tk.BooleanVar()  # Variable for highlight checkbox
        self.import_missing = tk.BooleanVar()  # Variable for import missing checkbox
        
        # UI Setup
        self.setup_ui()

    def setup_ui(self):
        # Use a grid layout for better positioning

        # Upload DXF Button and Label (Top-left)
        self.dxf_button = tk.Button(self, text="Upload DXF", command=self.upload_dxf)
        self.dxf_button.grid(row=0, column=0, padx=20, pady=10, sticky='w')

        self.dxf_label = tk.Label(self, text="No DXF file uploaded")
        self.dxf_label.grid(row=1, column=0, padx=20, pady=5, sticky='w')

        # Upload Template Button and Label (Top-right)
        self.template_button = tk.Button(self, text="Upload Template Excel", command=self.upload_template)
        self.template_button.grid(row=0, column=1, padx=20, pady=10, sticky='e')

        self.template_label = tk.Label(self, text="No Template Excel uploaded")
        self.template_label.grid(row=1, column=1, padx=20, pady=5, sticky='e')

        # Extract BOM Button (Center)
        self.extract_button = tk.Button(self, text="Extract BOM to Excel", state=tk.DISABLED, command=self.extract_bom)
        self.extract_button.grid(row=2, column=0, columnspan=2, padx=20, pady=20)

        # Upload Revised Excel Button and Label (Center)
        self.revised_button = tk.Button(self, text="Upload Revised BOM Excel", command=self.upload_revised_bom)
        self.revised_button.grid(row=3, column=0, padx=20, pady=10, sticky='w')

        self.revised_label = tk.Label(self, text="No Revised BOM uploaded")
        self.revised_label.grid(row=4, column=0, padx=20, pady=10, sticky='w')

        # Checkbox to select whether to highlight missing components (Center)
        self.highlight_checkbox = tk.Checkbutton(self, text="Highlight missing components in red", variable=self.highlight_missing)
        self.highlight_checkbox.grid(row=3, column=1, padx=20, pady=10, sticky='e')

        # Checkbox to select whether to import missing DXF items into Excel (Center)
        self.import_checkbox = tk.Checkbutton(self, text="Import missing DXF items to Excel", variable=self.import_missing)
        self.import_checkbox.grid(row=4, column=1, padx=20, pady=10, sticky='e')

        # Button to Compare BOM (Bottom-Center)
        self.compare_button = tk.Button(self, text="Compare BOM vs DXF", state=tk.DISABLED, command=self.compare_bom)
        self.compare_button.grid(row=5, column=0, columnspan=2, padx=20, pady=20)

    def upload_dxf(self):
        # Open file dialog to select a DXF file
        self.dwg_file = filedialog.askopenfilename(filetypes=[("DXF files", "*.dxf")])
        if self.dwg_file:
            logging.info(f"Uploaded DXF file: {self.dwg_file}")
            self.dxf_label.config(text=os.path.basename(self.dwg_file))
            self.check_ready_to_extract()

    def upload_template(self):
        # Open file dialog to select a template Excel file
        self.template_BOM_xls_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.template_BOM_xls_path:
            logging.info(f"Uploaded Excel template: {self.template_BOM_xls_path}")
            self.template_label.config(text=os.path.basename(self.template_BOM_xls_path))
            self.check_ready_to_extract()

    def upload_revised_bom(self):
        # Open file dialog to select a revised BOM Excel file
        self.revised_excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.revised_excel_file:
            logging.info(f"Uploaded Revised BOM Excel file: {self.revised_excel_file}")
            self.revised_label.config(text=os.path.basename(self.revised_excel_file))
            self.compare_button.config(state=tk.NORMAL)

    def check_ready_to_extract(self):
        # Enable "Extract BOM" button if both DXF and Template Excel are uploaded
        if self.dwg_file and self.template_BOM_xls_path:
            self.extract_button.config(state=tk.NORMAL)

    def extract_bom(self):
        try:
            logging.info("Extracting BOM from DXF...")
            self.bom_dxf = extract_bom_from_dxf(self.dwg_file)
            output_path = os.path.splitext(self.dwg_file)[0] + "_bom.xlsx"
            export_bom_to_excel(self.bom_dxf, self.template_BOM_xls_path, output_path)
            messagebox.showinfo("Success", f"BOM successfully exported to {output_path}")
        except Exception as e:
            logging.error(f"An error occurred during BOM extraction: {e}")
            messagebox.showerror("Error", f"Failed to extract BOM: {e}")

    def compare_bom(self):
        try:
            logging.info("Comparing BOM with DXF...")
            if self.bom_dxf is None:
                logging.error("BOM has not been extracted yet.")
                messagebox.showerror("Error", "BOM must be extracted before comparison.")
                return

            # Load the revised Excel file as a DataFrame
            bom_revised = load_bom_from_excel(self.revised_excel_file)

            # Convert BOM from DXF to DataFrame
            bom_df_dxf = convert_bom_dxf_to_dataframe(self.bom_dxf)

            # Check if the user wants to highlight missing components
            highlight_missing = self.highlight_missing.get()
            
            # Check if the user wants to import missing items from DXF to Excel
            import_missingDXF2BOM = self.import_missing.get()

            # Perform BOM comparison
            missing_in_revised, missing_in_dxf = compare_boms(bom_df_dxf, bom_revised, self.revised_excel_file, highlight_missing, import_missingDXF2BOM)
            
            # Display comparison results
            messagebox.showinfo("Comparison Results", f"Missing in Revised: {missing_in_revised}\nMissing in DXF: {missing_in_dxf}")
        except Exception as e:
            logging.error(f"An error occurred during BOM comparison: {e}")
            messagebox.showerror("Error", f"Failed to compare BOM: {e}")

if __name__ == "__main__":
    app = BOMExtractorApp()
    app.mainloop()
