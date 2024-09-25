import tkinter as tk
from tkinter import filedialog, messagebox
import logging
# Import the pyPidBoMExtractor package
from pyPidBoMExtractor.bom_generator import export_bom_to_excel, extract_bom_from_dxf
from pyPidBoMExtractor.bom_generator import compare_bomsJSON,convert_bom_dxf_to_JSON,load_bom_from_excel_to_JSON
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
        self.flagSaveNewExcellFile = tk.BooleanVar(value=True) 
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

        # # Extract BOM Button (Center)
        # self.extract_button = tk.Button(self, text="Extract BOM to Excel", state=tk.DISABLED, command=self.extract_bom)
        # self.extract_button.grid(row=2, column=0, columnspan=2, padx=20, pady=20)
        # Extract BOM Button
        self.extract_button = tk.Button(self, text="Extract BOM from DXF", state=tk.DISABLED, command=self.extract_bom)
        self.extract_button.grid(row=2, column=0, columnspan=2, padx=20, pady=10,  sticky='w') 
        
        # Export BOM to Excel Button (initially disabled)
        self.export_button = tk.Button(self, text="Export to Excel", state=tk.DISABLED, command=self.export_to_excel)
        self.export_button.grid(row=2, column=1, columnspan=2, padx=20, pady=10,  sticky='e')

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
        
        # Checkbox to select whether save new file or modifiy exisiting file
        self.import_checkbox = tk.Checkbutton(self, text="Save as new updated excell File", variable=self.flagSaveNewExcellFile)
        self.import_checkbox.grid(row=5, column=1, padx=20, pady=10, sticky='e')

        # Button to Compare BOM (Bottom-Center)
        self.compare_button = tk.Button(self, text="Compare BOM vs DXF", state=tk.DISABLED, command=self.compare_bom)
        self.compare_button.grid(row=6, column=0, columnspan=2, padx=20, pady=20)

    def upload_dxf(self):
        self.dwg_file = filedialog.askopenfilename(filetypes=[("DXF files", "*.dxf")])
        if self.dwg_file:
            self.dxf_label.config(text=os.path.basename(self.dwg_file))
            logging.info(f"Uploaded DXF file: {self.dwg_file}")
            self.extract_button.config(state=tk.NORMAL)

    def upload_template(self):
        self.template_BOM_xls_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.template_BOM_xls_path:
            self.template_label.config(text=os.path.basename(self.template_BOM_xls_path))
            logging.info(f"Uploaded Excel template: {self.template_BOM_xls_path}")
            if self.bom_dxf:
                self.export_button.config(state=tk.NORMAL)

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
           if not self.dwg_file:
               messagebox.showerror("Error", "Please upload a DXF file first.")
               return
        
           logging.info("Extracting BOM from DXF...")
           self.bom_dxf = extract_bom_from_dxf(self.dwg_file)
        
           # Enable the Export button once BOM is extracted
           self.export_button.config(state=tk.NORMAL)

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
                export_bom_to_excel(self.bom_dxf, self.template_BOM_xls_path, output_path)
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
            bom_revisedJSON = load_bom_from_excel_to_JSON(self.revised_excel_file)
    
            # Convert BOM from DXF to JSON
            bom_dxf = convert_bom_dxf_to_JSON(self.bom_dxf)
    
            # Check if the user wants to highlight missing components
            highlight_missing = self.highlight_missing.get()
            
            # Check if the user wants to import missing items from DXF to Excel
            import_missingDXF2BOM = self.import_missing.get()
            flagSaveNewExcellFile = self.flagSaveNewExcellFile.get()
    
            # If the "Save As New" checkbox is ticked, prompt for a new file name

            output_path = self.revised_excel_file
            
            # Check if the directory exists; if not, create it
            output_dir = os.path.dirname(output_path)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                logging.info(f"Created directory: {output_dir}")
    
            # Perform BOM comparison
            missing_in_revised, missing_in_dxf, workbook_excel = compare_bomsJSON(
                bom_dxf,
                bom_revisedJSON,
                revised_excel_file=output_path,  # Save to the selected path
                highlight_missing=highlight_missing,
                import_missingDXF2BOM=import_missingDXF2BOM
            )
            
            # Display comparison results
            messagebox.showinfo("Comparison Results", f"Missing in Revised: {missing_in_revised}\nMissing in DXF: {missing_in_dxf}")

            if flagSaveNewExcellFile:
                rev_excel_file_name = os.path.basename(self.revised_excel_file)
                default_filename = os.path.splitext(rev_excel_file_name)[0] + "_updated.xlsx"
                output_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    initialfile=default_filename,
                    title="Save updated BOM as"
                )
                workbook_excel.save(output_path)  # Save changes to the Excel file
                if not output_path:
                    messagebox.showerror("Error", "No file selected for saving the updated BOM.")
                    return
            else:
                workbook_excel.save(output_path)  # Save changes to the Excel file
                
           
        except Exception as e:
            logging.error(f"An error occurred during BOM comparison: {e}")
            messagebox.showerror("Error", f"Failed to compare BOM: {e}")


if __name__ == "__main__":
    app = BOMExtractorApp()
    app.mainloop()
