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
from pyPidBoMExtractor.bom_generator import generate_bom, print_bom, export_bom_to_excel

dwg_file = r"cad/Schema di funzionamento_rev1.1.dxf"
components = extract_blocks_with_attributes_and_dimensions(dwg_file)

# Print the formatted BOM with dimensions
tagBlocks = [c for c in components if isTagBlock(c)]
bom = generate_bom(components)
print_bom(bom)

templates_path = r"../templates/"
template_BOM_xls_path = templates_path+"BOM_ULIX_template.xlsx"
output_path = r"xls/BOM_1.xlsx"
export_bom_to_excel(bom, template_BOM_xls_path, output_path)