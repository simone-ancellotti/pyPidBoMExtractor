# -*- coding: utf-8 -*-
"""
Created on Fri Sep 20 15:50:07 2024

@author: user
"""

import argparse
import os
import sys
# Get the absolute path of the parent directory of pyPidBoMExtractor
# package_path = os.path.abspath('../pyPidBoMExtractor')

# # Insert the package path into sys.path
# sys.path.insert(0, package_path)

# Import the pyPidBoMExtractor package
import pyPidBoMExtractor
from pyPidBoMExtractor.extractor import extract_blocks_with_attributes_and_dimensions, isTagBlock
from pyPidBoMExtractor.bom_generator import generate_bom, print_bom, export_bom_to_excel
from pyPidBoMExtractor import bom_generator  # Assuming bom_generator has the logic to generate BOM from DXF

def main():
    # Set up argument parsing
    parser = argparse.ArgumentParser(description='Generate BOM from a DXF file.')
    
    # Argument for the input DXF file
    parser.add_argument('-d', '--dxf', required=True, help='Path to the input DXF file')
    
    # Optional argument for the output Excel file
    parser.add_argument('-o', '--output', help='Path to the output Excel file (optional). If not provided, it will use the input DXF file name for output.')

    parser.add_argument('-t', '--template', help='Path to the template Excel file (optional). If not provided, it will use the default file name as tempplate.')

    # Parse the arguments
    args = parser.parse_args()

    # Get the input DXF file path
    input_dxf = args.dxf

    # If the output path is not provided, generate a default name
    if args.output:
        output_excel = args.output
    else:
        # Use the base name of the input DXF file, replacing the extension with "_bom.xlsx"
        base_name = os.path.splitext(os.path.basename(input_dxf))[0]
        output_excel = f"{base_name}_bom.xlsx"
    

    if args.template:
        template_BOM_xls_path = args.template
    else:
        templates_path = r"templates/"
        template_BOM_xls_path = templates_path+"BOM_ULIX_template.xlsx"
    # Now call your function to generate the BOM using the input DXF and output Excel
    # bom_data = bom_generator.extract_bom_from_dxf(input_dxf)  # Assuming this is the function that extracts the BOM data
    # bom_generator.export_bom_to_excel(bom_data, 'templates/BOM_ULIX_template.xlsx', output_excel)
    
    components = extract_blocks_with_attributes_and_dimensions(input_dxf)

    # Print the formatted BOM with dimensions
    bom = generate_bom(components)
    print_bom(bom)


    #output_path = r"xls/BOM_1.xlsx"
    export_bom_to_excel(bom, template_BOM_xls_path, output_excel)
    
    print(f"BOM generated and saved to {output_excel}")

if __name__ == "__main__":
    main()
