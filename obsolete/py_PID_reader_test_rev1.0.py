import ezdxf

def extract_blocks_with_attributes(dwg_file_path):
    # Load the DWG file
    dwg = ezdxf.readfile(dwg_file_path)
    msp = dwg.modelspace()

    components = []

    # Iterate over all block references (INSERT entities) in the modelspace
    for entity in msp.query('INSERT'):
        block_name = entity.dxf.name
        insert_point = entity.dxf.insert

        # Extract attributes associated with the block
        attributes = {}
        for attrib in entity.attribs:
            attributes[attrib.dxf.tag] = attrib.dxf.text

        # Check if the block definition contains a circle
        block_def = dwg.blocks.get(block_name)
        contains_circle = any(e.dxftype() == 'ARC' for e in block_def)

        component = {
            "block_name": block_name,
            "insert_point": (insert_point.x, insert_point.y),
            "attributes": attributes,
            "contains_circle": contains_circle
        }
        components.append(component)

    return components

def generate_bom(components):
    bom = {}
    for component in components:
        block_name = component['block_name']
        attributes = component['attributes']
        contains_circle = component['contains_circle']

        if 'NUMBER' in attributes:
            number = attributes['NUMBER']
            if number in bom:
                bom[number]['count'] += 1
            else:
                bom[number] = {"count": 1, "contains_circle": contains_circle}
        else:
            print(f"Block {block_name} does not have a 'NUMBER' attribute.")

    return bom

if __name__ == "__main__":
    dwg_file = r"P&ID_simple.dxf"  # Corrected the file path
    
    components = extract_blocks_with_attributes(dwg_file)
    bom = generate_bom(components)
    
    # Print the BOM with circle information
    for component, info in bom.items():
        print(f"Component {component}: {info['count']} instances, Contains Circle: {info['contains_circle']}")

    print(' \n')
    for comp in components:
        print(comp)