import sys
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)

# Print the current working directory
logging.info(f"Current working directory: {os.getcwd()}")

# Get the absolute path of the parent directory of pyPidBoMExtractor
package_path = os.path.abspath('../')

# Insert the package path into sys.path
sys.path.insert(0, package_path)

# Import the pyPidBoMExtractor package
from pyPidBoMExtractor.bom_generator import export_bom_to_excel, extract_bom_from_dxf
from pyPidBoMExtractor.bom_generator import compare_bomsJSON,convert_bom_dxf_to_JSON,load_bom_from_excel_to_JSON
from pyPidBoMExtractor.bom_generator import sortingBOM_dict

dwg_file = r"cad/Schema di funzionamento_rev1.1.dxf"
templates_path = r"../templates/"
template_BOM_xls_path = os.path.join(templates_path, "BOM_ULIX_template.xlsx")
output_path = r"xls/BOM_1.xlsx"
revised_excel_file = r"xls/BOM_1.1.xlsx"
revised_excel_file2 = r"xls/BOM_1.1_updated.xlsx"


if __name__ == "__main__":
    try:
        # Extract BOM from DXF
        logging.info("Extracting BOM from DXF...")
        bom_dxf = extract_bom_from_dxf(dwg_file)
        # bom_dxf = sortingBOM_dict(bom_dxf,bytag = 'P&ID TAG' )
        
        export_bom_to_excel(bom_dxf, template_BOM_xls_path, output_path)
        logging.info(f"BOM successfully exported to {output_path}")

        # Load BOM from revised Excel file
        logging.info("Loading revised BOM from Excel...")
        # bom_revised = load_bom_from_excel(revised_excel_file)
        bom_revisedJSON = load_bom_from_excel_to_JSON(revised_excel_file)
        #bom_revisedJSON = sortingBOM_dict(bom_revisedJSON,bytag = 'P&ID TAG' )
        # Convert DXF BOM to DataFrame
        # bom_df_dxf = convert_bom_dxf_to_dataframe(bom_dxf)
        
        bom_dxf_updated_keys = convert_bom_dxf_to_JSON(bom_dxf)
        

        # Ask user if they want to highlight missing components
        # highlight_option = input("Do you want to mark missing components in red in the revised Excel file? <yes,no>: ").strip().lower()
        # highlight_missing = highlight_option == 'yes'
        highlight_missing = True
        
        # import_missing = input("Do you want to import missing items from DXF to Excel? New added items rows will be highlighted in grey. <yes,no>: ").strip().lower()
        # import_missingDXF2BOM = import_missing == 'yes'
        import_missingDXF2BOM = True
        
        flagSaveNewExcellFile = True
        # Compare BOMs
        #missing_in_revised, missing_in_dxf = compare_boms(bom_df_dxf, bom_revised, revised_excel_file, highlight_missing,import_missingDXF2BOM)
        missing_in_revised, missing_in_dxf, workbook_excel= compare_bomsJSON(bom_dxf_updated_keys, bom_revisedJSON, revised_excel_file, highlight_missing,import_missingDXF2BOM)
        
        workbook_excel.save(revised_excel_file2)  # Save changes to the Excel file
    except FileNotFoundError as e:
        logging.error(f"File not found: {e}")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
