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
import pyPidBoMExtractor
from pyPidBoMExtractor.extractor import extract_blocks_with_attributes_and_dimensions, isTagBlock
from pyPidBoMExtractor.bom_generator import generate_bom, print_bom, export_bom_to_excel, extract_bom_from_dxf, load_bom_from_excel
from pyPidBoMExtractor.bom_generator import compare_boms, convert_bom_dxf_to_dataframe

dwg_file = r"cad/Schema di funzionamento_rev1.1.dxf"
templates_path = r"../templates/"
template_BOM_xls_path = os.path.join(templates_path, "BOM_ULIX_template.xlsx")
output_path = r"xls/BOM_1.xlsx"
revised_excel_file = r"xls/BOM_1.1.xlsx"

def main():
    try:
        # Extract BOM from DXF
        logging.info("Extracting BOM from DXF...")
        bom_dxf = extract_bom_from_dxf(dwg_file)
        export_bom_to_excel(bom_dxf, template_BOM_xls_path, output_path)
        logging.info(f"BOM successfully exported to {output_path}")

        # Load BOM from revised Excel file
        logging.info("Loading revised BOM from Excel...")
        bom_revised = load_bom_from_excel(revised_excel_file)

        # Convert DXF BOM to DataFrame
        bom_df_dxf = convert_bom_dxf_to_dataframe(bom_dxf)

        # Ask user if they want to highlight missing components
        highlight_option = input("Do you want to mark missing components in red in the revised Excel file? <yes,no>: ").strip().lower()
        highlight_missing = highlight_option == 'yes'

        # Compare BOMs
        missing_in_revised, missing_in_dxf = compare_boms(bom_df_dxf, bom_revised, revised_excel_file, highlight_missing)

    except FileNotFoundError as e:
        logging.error(f"File not found: {e}")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
