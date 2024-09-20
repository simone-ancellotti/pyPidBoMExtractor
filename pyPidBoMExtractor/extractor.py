import ezdxf
from .utils import get_block_final_bounding_box, isTagBlock


def extract_blocks_with_attributes_and_dimensions(dwg_file_path):
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

        # Get the scaling factors
        scale_x = getattr(entity.dxf, 'xscale', 1.0)  # Default to 1.0 if not found
        scale_y = getattr(entity.dxf, 'yscale', 1.0)  # Default to 1.0 if not found
        scale_z = getattr(entity.dxf, 'zscale', 1.0)  # Default to 1.0 if not found

        # Check if the block definition contains a circle
        block_def = dwg.blocks.get(block_name)
        contains_circle = any(e.dxftype() == 'CIRCLE' for e in block_def)

        # Print the block entities for debugging
        #print(f"Block: {block_name}")
        #print_block_entities(block_def)
        
        # Calculate the final bounding box dimensions of the block
        dimensions = get_block_final_bounding_box(block_def, scale_x, scale_y, scale_z)
        
        if dimensions is None:
            # Assign default dimensions if no geometric entities were found
            dimensions = {"width": 10 * scale_x, "height": 10 * scale_y}

        component = {
            "block_name": block_name,
            "insert_point": (insert_point.x, insert_point.y),
            "attributes": attributes,
            "contains_circle": contains_circle,
            "dimensions": dimensions,  # Adding dimensions to the component
            "block_def" : block_def
        }
        components.append(component)

    return components
