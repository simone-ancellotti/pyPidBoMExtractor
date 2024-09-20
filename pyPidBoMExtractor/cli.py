import argparse
from .extractor import extract_blocks_with_attributes_and_dimensions
from .bom_generator import generate_bom, print_bom

def main():
    parser = argparse.ArgumentParser(description="Generate BOM from DXF P&ID")
    parser.add_argument('file', help='Path to the DXF file')
    args = parser.parse_args()

    components = extract_blocks_with_attributes_and_dimensions(args.file)
    bom = generate_bom(components)
    print_bom(bom)

if __name__ == "__main__":
    main()
