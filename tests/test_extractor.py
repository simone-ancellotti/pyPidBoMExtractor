import sys
import os

# Print the current working directory
print(os.getcwd())

# Get the absolute path of the parent directory of pyPidBoMExtractor
package_path = os.path.abspath('../pyPidBoMExtractor')

# Insert the package path into sys.path
sys.path.insert(0, package_path)

# Import the pyPidBoMExtractor package
import pyPidBoMExtractor
from pyPidBoMExtractor.extractor import extract_blocks_with_attributes_and_dimensions, isTagBlock
from pyPidBoMExtractor.bom_generator import generate_bom, print_bom, export_bom_to_excel, extract_bom_from_dxf, load_bom_from_excel
from  pyPidBoMExtractor.bom_generator import compare_boms, convert_bom_dxf_to_dataframe
dwg_file = r"cad/Schema di funzionamento_rev1.1.dxf"
templates_path = r"../templates/"
template_BOM_xls_path = templates_path+"BOM_ULIX_template.xlsx"
output_path = r"xls/BOM_1.xlsx"
revised_excel_file = r"xls/BOM_1.1.xlsx"

# Extract BOM from DXF (assuming this function exists)
bom_dxf = extract_bom_from_dxf(dwg_file)
export_bom_to_excel(bom_dxf, template_BOM_xls_path, output_path)


# # Load BOM from revised Excel file
bom_revised = load_bom_from_excel(revised_excel_file)

# # Compare BOMs
bom_df_dxf = convert_bom_dxf_to_dataframe(bom_dxf)
compare_boms(bom_df_dxf, bom_revised)