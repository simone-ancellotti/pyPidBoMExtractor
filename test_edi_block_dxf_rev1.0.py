# -*- coding: utf-8 -*-
"""
Created on Wed Apr  9 10:14:45 2025

@author: user
"""

from pyPidBoMExtractor.bom_generator import export_bom_to_excel, extract_bom_from_dxf,filterBOM_Ignore
from pyPidBoMExtractor.bom_generator import compare_bomsJSON,convert_bom_dxf_to_JSON,load_bom_from_excel_to_JSON
import ezdxf

def get_attrib_from_tag(tag, block):
    """
    Retrieve an attribute (ATTDEF/ATTRIB) from a block by its P&ID tag.
    
    The function works for two block types:
      - ezdxf.entities.insert.Insert: uses 'attribs'
      - ezdxf.layouts.blocklayout.BlockLayout: uses 'attdefs()'
    
    Args:
        tag (str): The P&ID tag to search for.
        block: The block entity from which to retrieve the attribute.
        
    Returns:
        The attribute object if found; otherwise, None.
        
    Raises:
        ValueError: If the block type is not supported.
    """
    
    # Normalize tag for case-insensitive comparison.
    normalized_tag = tag.upper()

    # Determine how to obtain the attributes based on block type
    if isinstance(block, ezdxf.entities.insert.Insert):
        # For Insert type blocks, use the 'attribs' attribute (already iterable)
        list_of_attribs = block.attribs
    elif isinstance(block, ezdxf.layouts.blocklayout.BlockLayout):
        # For BlockLayout blocks, use the 'attdefs()' method.
        list_of_attribs = block.attdefs()
    else:
        raise ValueError(f"Unsupported block type: {type(block)}")
    
    # Iterate over the available attributes and compare the tags (case-insensitive).
    for attrib in list_of_attribs:
        if attrib.dxf.tag.upper() == normalized_tag:
            return attrib
    return None



def add_new_tag_to_insert(tag, text='', insert_block=None):
    """
    Add a new attribute/tag to a block of type ezdxf.entities.insert.Insert.
    
    The function performs the following steps:
      1. Checks if the tag already exists in the insert block.
      2. If not, adds a new attribute to the insert block using add_attrib().
      3. Updates the corresponding block definition so that the attribute is
         synchronized for future insertions by calling add_attdef().
    
    Args:
        tag (str): The attribute tag to be added (case-insensitive).
        text (str, optional): The default text value for the new attribute.
                              Defaults to '' which is replaced with 'default_value'.
        insert_block (ezdxf.entities.insert.Insert): The block reference to update.
    
    Returns:
        new_attrib: The newly added attribute from the insert block. If the attribute
                    already exists, the existing attribute is returned.
    
    Raises:
        ValueError: If the provided block is not an instance of ezdxf.entities.insert.Insert.
    """
    if insert_block is None:
        raise ValueError("You must provide an insert block.")

    # Ensure the insert_block is the correct type.
    if not isinstance(insert_block, ezdxf.entities.insert.Insert):
        raise ValueError("The provided block must be of type ezdxf.entities.insert.Insert.")

    # Set a default text value if text parameter is empty.
    text_val = text if text else 'default_value'

    # Check if the attribute already exists (case-insensitive check).
    for attrib in insert_block.attribs:
        if attrib.dxf.tag.upper() == tag.upper():
            print(f"Attribute with tag '{tag}' already exists in the insert.")
            return attrib

    # Step 2: Add a new attribute to the insert block.
    new_attrib = insert_block.add_attrib(
        tag=tag,
        text=text_val,
        insert=(0, 0),  # Adjust the insertion point as needed.
        dxfattribs={'flags': 1}
    )
    print(f"Added new attribute '{tag}' to the insert block.")

    # Step 3: Synchronize with the block definition.
    # Retrieve the document from the insert block.
    doc = insert_block.doc
    # The block name is stored in the dxf.name attribute.
    block_name = insert_block.dxf.name
    block_def = doc.blocks.get(block_name)

    # Check if the attribute definition already exists in the block definition.
    for attdef in block_def.attdefs():
        if attdef.dxf.tag.upper() == tag.upper():
            print(f"Attribute definition '{tag}' already exists in block definition '{block_name}'.")
            break
    else:
        # Add new attribute definition to the block definition.
        new_attrib = block_def.add_attdef(
            tag=tag,
            insert=(0, 0),  # Relative insertion point inside the block
            text='',
            height=2.5,     # Default text height; adjust as needed.
            rotation=0,
            dxfattribs={'flags': 1}
        )
        print(f"Added new attribute definition '{tag}' to block definition '{block_name}'.")

    return new_attrib



def update_tag_value_in_block(text, tag, insert_block):
    """
    Update the attribute value for a given tag in an ezdxf.entities.insert.Insert block.
    
    If the attribute identified by 'tag' already exists within the insert block,
    its 'text' value will be updated. Otherwise, the attribute will be created 
    (and synchronized with the block definition) with the provided text value.
    
    Args:
        text (str): The new text to assign to the attribute. If empty, 'default_value' is used.
        tag (str): The attribute tag identifying the attribute (case insensitive).
        insert_block (ezdxf.entities.insert.Insert): The DXF block instance to update.
        
    Returns:
        The updated or newly created attribute object.
        
    Raises:
        ValueError: If the insert_block is None or not an instance of ezdxf.entities.insert.Insert.
    """
    if insert_block is None:
        raise ValueError("You must provide an insert block.")
    
    if not isinstance(insert_block, ezdxf.entities.insert.Insert):
        raise ValueError("The provided block must be of type ezdxf.entities.insert.Insert.")
    
    # Use the provided text or default to 'default_value'
    text_val = text if text else 'default_value'
    
    # Check for an existing attribute with the specified tag.
    attrib = get_attrib_from_tag(tag, insert_block)
    
    if attrib:
        # If the attribute exists, update its text value.
        attrib.dxf.text = text_val
        return attrib
    else:
        # If the attribute does not exist, create a new one and synchronize with the block definition.
        new_attrib = add_new_tag_to_insert(tag, text=text_val, insert_block=insert_block)
        return new_attrib


listTagBlockNames = ('Tag_block','Ball_Tag','Ball_Tag2',
                     'DisplayControl0_Tag','DisplayControl1_Tag','DisplayControl2_Tag',
                     'LogicControl0_Tag','LogicControl1_Tag','LogicControl2_Tag',
                     'STICKER Moving Machine', 'STICKER Equipment Name',
                     'Tag_Instrument')


# 1. Load your DXF file
doc = ezdxf.readfile("P&ID_simple_template_test_import.dxf")

for block_names in doc.blocks.block_names():
    block_names 

# 2. Get the block definition by name (replace 'MY_BLOCK' with your block name)

block_def = doc.blocks.get('psv')

# 3. Add an attribute definition to the block definition

#    The add_attdef() method creates an ATTDEF entity in the block definition.

new_attdef = block_def.add_attdef(
    tag="NEW_TAG",        # The new attribute tag (i.e. "NEW_TAG")
    insert=(0, 0),        # Insertion point relative to the block's coordinate system
    text="default_value", # Default text for the attribute
    height=2.5,           # Text height
    rotation=0,
    dxfattribs = {'flags': 1}
)


#clone_attribute_with_new_tag(block_def, 'TYPE', 'new_tag')

for attdef in  block_def.attdefs():
     print(attdef.dxfattribs())


    #attdef.transparency=0
print("Added new attribute definition with tag 'NEW_TAG' to the block.")
block_def.update_block_flags()

# 4. For each block reference (INSERT entity) that uses 'MY_BLOCK', add an attribute reference
msp = doc.modelspace()
for insert in msp.query("INSERT[name=='PSV']"):
    # Check if the attribute already exists (optional)
    if not any(att.dxf.tag.upper() == "NEW_TAG" for att in insert.attribs):
        # Add a new attribute to the INSERT entity using add_attrib()
        new_attrib = insert.add_attrib(
            tag="NEW_TAG",        # Tag should match the ATTDEF
            text="default_value", # You can set an initial value here
            insert=(0, 0),         # Position relative to the insert; adjust as needed
            dxfattribs = {'flags': 1}
        )
        print(f"Added attribute reference 'NEW_TAG' to block INSERT at {insert.dxf.insert}.")
        add_new_tag_to_insert("NEW_TAG2", text='', insert_block=insert)
        update_tag_value_in_block(text='simone', tag="NEW_TAG2", insert_block=insert)
    # Now update block references to reflect changes:
        




get_attrib_from_tag('type',block_def)
get_attrib_from_tag('type',insert)
 

# 5. Save the modified DXF file.
doc.saveas("my_drawing_modified.dxf")
print("Modified DXF saved as 'my_drawing_modified.dxf'.")