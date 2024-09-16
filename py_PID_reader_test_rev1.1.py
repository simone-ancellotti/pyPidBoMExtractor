import ezdxf

def is_hexagon(entities):
    """
    Check if any of the entities is a hexagon (a polyline with 6 vertices).
    """
    for entity in entities:
        if entity.dxftype() == 'LWPOLYLINE':
            # For LWPOLYLINE entities
            if hasattr(entity, 'points'):
                points = list(entity.points())
                if len(points) == 6:
                    if entity.is_closed:
                        return True
        elif entity.dxftype() == 'POLYLINE':
            # For POLYLINE entities
            if hasattr(entity, 'vertices'):
                points = list(entity.vertices())
                if len(points) == 6:
                    if entity.is_closed:
                        return True
    return False

def extract_blocks_with_attributes(dwg_file_path):
    # Load the DXF file
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

        # Check if the block definition contains a hexagon
        block_def = dwg.blocks.get(block_name)
        contains_hexagon = is_hexagon(block_def)

        component = {
            "block_name": block_name,
            "insert_point": (insert_point.x, insert_point.y),
            "attributes": attributes,
            "contains_hexagon": contains_hexagon
        }
        components.append(component)

    return components

def generate_bom(components):
    bom = {}
    for component in components:
        block_name = component['block_name']
        attributes = component['attributes']
        contains_hexagon = component['contains_hexagon']

        if 'NUMBER' in attributes:
            number = attributes['NUMBER']
            if number in bom:
                bom[number]['count'] += 1
            else:
                bom[number] = {"count": 1, "contains_hexagon": contains_hexagon}
        else:
            print(f"Block {block_name} does not have a 'NUMBER' attribute.")

    return bom

if __name__ == "__main__":
    dwg_file = r"P&ID_simple.dxf"  # Your DXF file path
    
    components = extract_blocks_with_attributes(dwg_file)
    bom = generate_bom(components)
    
    # Print the BOM with hexagon information
    for component, info in bom.items():
        print(f"Component {component}: {info['count']} instances, Contains Hexagon: {info['contains_hexagon']}")
