# -*- coding: utf-8 -*-
"""
Created on Wed Apr  9 10:14:45 2025

@author: user
"""

import ezdxf

def clone_attribute_with_new_tag(block_def, existing_tag, new_tag):
    """
    Clone an existing attribute definition (ATTDEF) from block_def that has the given existing_tag,
    and create a new attribute definition with a new tag. This new attribute
    should inherit the properties (including flags) from the original.
    
    Args:
        block_def: The block definition (BlockTableRecord) where the attribute resides.
        existing_tag: The tag value of the existing attribute to copy (e.g. "OLD_TAG").
        new_tag: The new tag value for the cloned attribute (e.g. "NEW_TAG").
        
    Returns:
        The newly created attribute definition.
        
    Raises:
        ValueError: If no attribute with the specified existing_tag is found.
    """
    # Loop through existing attribute definitions in the block
    for attdef in block_def.attdefs():
        if attdef.dxf.tag.upper() == existing_tag.upper():
            # Get dictionary of DXF attributes, excluding values that are automatically assigned
            props = attdef.dxfattribs(exclude=["handle", "owner_handle"])
            # Update the tag to the new value
            props["tag"] = new_tag
            # Create a new attribute definition in the block using these properties
            new_attdef = block_def.add_attdef(**props)
            return new_attdef
    raise ValueError(f"No attribute with tag '{existing_tag}' found in the block definition.")


# 1. Load your DXF file
doc = ezdxf.readfile("P&ID_simple_template_test_import.dxf")


# 2. Get the block definition by name (replace 'MY_BLOCK' with your block name)
doc.blocks.block_names()
block_def = doc.blocks.get('psv')

# 3. Add an attribute definition to the block definition

#    The add_attdef() method creates an ATTDEF entity in the block definition.

new_attdef = block_def.add_attdef(
    tag="NEW_TAG",        # The new attribute tag (i.e. "NEW_TAG")
    insert=(0, 0),        # Insertion point relative to the block's coordinate system
    text="default_value", # Default text for the attribute
    height=2.5,           # Text height
    rotation=0,  
)


for attdef in  block_def.attdefs():
    print(attdef.dxfattribs())

    #attdef.transparency=0
print("Added new attribute definition with tag 'NEW_TAG' to the block.")
block_def.update_block_flags()

# 4. For each block reference (INSERT entity) that uses 'MY_BLOCK', add an attribute reference
msp = doc.modelspace()
for insert in msp.query("INSERT[name=='psv']"):
    # Check if the attribute already exists (optional)
    if not any(att.dxf.tag.upper() == "NEW_TAG" for att in insert.attribs):
        # Add a new attribute to the INSERT entity using add_attrib()
        new_attrib = insert.add_attrib(
            tag="NEW_TAG",        # Tag should match the ATTDEF
            text="default_value", # You can set an initial value here
            insert=(0, 0)         # Position relative to the insert; adjust as needed
        )
        print(f"Added attribute reference 'NEW_TAG' to block INSERT at {insert.dxf.insert}.")
    # Now update block references to reflect changes:
        
    insert.attribs.sync()

# 5. Save the modified DXF file.
doc.saveas("my_drawing_modified.dxf")
print("Modified DXF saved as 'my_drawing_modified.dxf'.")