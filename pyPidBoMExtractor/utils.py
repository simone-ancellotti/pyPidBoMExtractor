#import numpy as np
import ezdxf
from ezdxf.math import BoundingBox, Vec3
import re
import math

listTagBlockNames = ('Tag_block','Ball_Tag','Ball_Tag2',
                     'DisplayControl0_Tag','DisplayControl1_Tag','DisplayControl2_Tag',
                     'LogicControl0_Tag','LogicControl1_Tag','LogicControl2_Tag',
                     'STICKER Moving Machine', 'STICKER Equipment Name',
                     'Tag_Instrument')


blocks_ignore_distance = ['ULIX_title_block']


def normalize(v):
    # Calculate the Euclidean norm (L2 norm)
    norm = math.sqrt(sum(x ** 2 for x in v))
    
    if norm == 0:
        return v  # Return the vector unchanged if the norm is 0
    
    # Divide each component by the norm to normalize the vector
    return [x / norm for x in v]

# def normalize(v):
#     v=np.array(v)
#     norm = np.linalg.norm(v)
#     if norm == 0: return v
#     return v / norm

def print_block_entities(block_def):
    """Print the details of each entity in a block definition."""
    for entity in block_def:
        print(f"Entity Type: {entity.dxftype()}")
        for attr in dir(entity):
            if not attr.startswith('_'):
                try:
                    value = getattr(entity, attr)
                    if value:
                        print(f"  {attr}: {value}")
                except Exception as e:
                    print(f"  {attr}: Error retrieving attribute - {e}")
        print("\n")


def get_block_dimensions(block_def):
    """Calculate the width and height of a block using its bounding box."""
    bounding_box = BoundingBox()

    for entity in block_def:
        # Handle different entity types for bounding box calculation
        if entity.dxftype() == 'LINE':
            start_point = entity.dxf.start
            end_point = entity.dxf.end
            bounding_box.extend([start_point, end_point])

        elif entity.dxftype() == 'CIRCLE':
            center = entity.dxf.center
            radius = entity.dxf.radius
            bounding_box.extend([
                (center.x - radius, center.y - radius, center.z),
                (center.x + radius, center.y + radius, center.z)
            ])

        elif entity.dxftype() == 'ARC':
            center = entity.dxf.center
            radius = entity.dxf.radius
            bounding_box.extend([
                (center.x - radius, center.y - radius, center.z),
                (center.x + radius, center.y + radius, center.z)
            ])

        elif entity.dxftype() == 'LWPOLYLINE' or entity.dxftype() == 'POLYLINE':
            for point in entity.get_points():
                bounding_box.extend([point[:3]])  # Assuming point has (x, y, z)
        
        # Handle TEXT and MTEXT entities for dimension calculation
        elif entity.dxftype() == 'TEXT':
            insert_point = entity.dxf.insert
            height = entity.dxf.height
            width_factor = entity.dxf.get("width_factor", 1.0)
            text_width = len(entity.text) * height * width_factor
            bounding_box.extend([
                (insert_point.x, insert_point.y, insert_point.z),
                (insert_point.x + text_width, insert_point.y + height, insert_point.z)
            ])
        
        elif entity.dxftype() == 'MTEXT':
            insert_point = entity.dxf.insert
            width = entity.dxf.get("width", 0)
            char_height = entity.dxf.char_height  # Use char_height directly for MTEXT
            bounding_box.extend([
                (insert_point.x, insert_point.y, insert_point.z),
                (insert_point.x + width, insert_point.y + char_height, insert_point.z)
            ])
        elif entity.dxftype() == 'ATTDEF':
        # ATTDEF entities do not have dimensions; they are just attribute definitions
            pass
        # Add other entity types if needed, like SPLINE, ELLIPSE, etc.

    if not bounding_box.has_data:
        # Return default dimensions if no geometric entities were found
        return {"width": 10, "height": 10}

    # Extract the minimum and maximum points from the bounding box
    min_point, max_point = bounding_box.extmin, bounding_box.extmax

    width = max_point.x - min_point.x
    height = max_point.y - min_point.y

    return {"width": width, "height": height}

def get_block_final_bounding_box(block_def, scale_x=1.0, scale_y=1.0, scale_z=1.0):
    """Calculate the final bounding box of a block, applying scaling transformations."""
    bounding_box = BoundingBox()

    for entity in block_def:
        # Handle different entity types for bounding box calculation
        if entity.dxftype() == 'LINE':
            start_point = entity.dxf.start
            end_point = entity.dxf.end
            bounding_box.extend([start_point, end_point])

        elif entity.dxftype() == 'CIRCLE':
            center = entity.dxf.center
            radius = entity.dxf.radius
            bounding_box.extend([
                (center.x - radius, center.y - radius, center.z),
                (center.x + radius, center.y + radius, center.z)
            ])

        elif entity.dxftype() == 'ARC':
            center = entity.dxf.center
            radius = entity.dxf.radius
            bounding_box.extend([
                (center.x - radius, center.y - radius, center.z),
                (center.x + radius, center.y + radius, center.z)
            ])

        elif entity.dxftype() == 'LWPOLYLINE':
            # For LWPOLYLINE, we can use the get_points method
            for point in entity.get_points():
                bounding_box.extend([point[:3]])  # Assuming point has (x, y, z)

        elif entity.dxftype() == 'POLYLINE':
            # For POLYLINE, use the vertices attribute to access points
            for vertex in entity.vertices:
                bounding_box.extend([vertex.dxf.location])

        elif entity.dxftype() == 'TEXT':
            insert_point = entity.dxf.insert
            height = entity.dxf.height
            width_factor = 1.0
            if entity.dxf.hasattr("width_factor"):
                width_factor = entity.dxf.get("width_factor", 1.0)
            text_width = len(entity.dxf.text) * height * width_factor
            bounding_box.extend([
                (insert_point.x, insert_point.y, insert_point.z),
                (insert_point.x + text_width, insert_point.y + height, insert_point.z)
            ])

        elif entity.dxftype() == 'MTEXT':
            insert_point = entity.dxf.insert
            width = entity.dxf.get("width", 0)
            char_height = entity.dxf.char_height  # Use char_height directly for MTEXT
            bounding_box.extend([
                (insert_point.x, insert_point.y, insert_point.z),
                (insert_point.x + width, insert_point.y + char_height, insert_point.z)
            ])

        elif entity.dxftype() == 'ATTDEF':
            # ATTDEF entities do not have dimensions; they are just attribute definitions
            pass

    if not bounding_box.has_data:
        return None  # No entities in the block or unable to calculate

    # Extract the minimum and maximum points from the bounding_box
    min_point, max_point = bounding_box.extmin, bounding_box.extmax

    # Create new points with scaling applied
    min_point_scaled = Vec3(min_point.x * scale_x, min_point.y * scale_y, min_point.z * scale_z)
    max_point_scaled = Vec3(max_point.x * scale_x, max_point.y * scale_y, max_point.z * scale_z)

    width = max_point_scaled.x - min_point_scaled.x
    height = max_point_scaled.y - min_point_scaled.y

    return {"width": width, "height": height}




def isTagBlock(block):
    block_name = block['block_name']
    if block_name[:2] == "*U" and block_name[2:].isdigit():
        return True
    if block_name in listTagBlockNames:
        return True
    return False

def getCaracteristicDimensionBlock(block):
    dimensionsComp = block['dimensions']
    return max(dimensionsComp['width'],dimensionsComp['height'])

def getCriticalDistance(dim1,dim2):
    return dim1/2.0 + dim2/2.0 + max(dim1,dim2) * 0.5

def calculate_distance(point1, point2):
    """
    Calculate the Euclidean distance between two points in any dimension.
    :param point1: List or tuple representing the coordinates of the first point (e.g., [x, y] or [x, y, z])
    :param point2: List or tuple representing the coordinates of the second point (e.g., [x, y] or [x, y, z])
    :return: Euclidean distance between the two points
    """
    return math.sqrt(sum((a - b) ** 2 for a, b in zip(point1, point2)))

def findBlocksNearBlock(block,components ):
    filtered_blocks = []
    distances_with_block0 = []
    # positionBlock0 = np.array(block['insert_point'])
    positionBlock0 = block['insert_point']
    mainDimBlock0  = getCaracteristicDimensionBlock(block)
    i=-1
    for component in components:
        i+=1
        flagIsToIgnore = component['block_name'] in blocks_ignore_distance
        # if flagIsToIgnore:
        #     print(component['block_name'])
        if block['block_def'] != component['block_def'] and not(flagIsToIgnore):
            # positionComponent1 = np.array(component['insert_point'])
            positionComponent1 = component['insert_point']
            mainDim1 = getCaracteristicDimensionBlock(component)
            
            criticalDistance = getCriticalDistance(mainDimBlock0,mainDim1)
            # distance = np.linalg.norm(positionBlock0 - positionComponent1)
            distance = calculate_distance(positionBlock0, positionComponent1)
            if distance <= criticalDistance:
                filtered_blocks.append(component)
                distances_with_block0.append(distance)
    
    # sort by distance
    sorted_filtered_blocks, sorted_distances = sort_blocks_by_distance(filtered_blocks, distances_with_block0)

    return sorted_filtered_blocks, sorted_distances

def sort_blocks_by_distance(filtered_blocks, distances_with_block0):
    """
    Sorts the filtered blocks based on their distances from the reference block.

    Args:
        filtered_blocks (list): List of blocks to be sorted.
        distances_with_block0 (list): Corresponding distances of the blocks.

    Returns:
        sorted_filtered_blocks (list): Blocks sorted by distance.
        sorted_distances (list): Distances sorted in ascending order.
    """
    # Combine filtered_blocks and distances_with_block0 into a list of tuples
    blocks_with_distances = list(zip(filtered_blocks, distances_with_block0))

    # Sort the list of tuples based on the second element (distance)
    blocks_with_distances.sort(key=lambda x: x[1])

    # Extract the sorted filtered_blocks and distances back
    sorted_filtered_blocks = [block for block, _ in blocks_with_distances]
    sorted_distances = [distance for _, distance in blocks_with_distances]

    return sorted_filtered_blocks, sorted_distances
      
def getTypeFromBlock(block):
    attributes = block['attributes']
    
    typeTag = None
    if 'TYPE' in attributes.keys():
        typeTag = attributes['TYPE']
    # else: 
    #     typeTag = attributes.get('DESCRIPTION')
    return typeTag

def findTypeBlockFromTag(block,components):
    #blocksValveAndMachine = [comp for comp in components ]
    # attributes = block['attributes']
    
    # typeTag = ''
    # if 'TYPE' in attributes.items():
    #     typeTag = attributes['TYPE']
    typeTag = getTypeFromBlock(block)
    distance = None
    block_found = None
    
    flag_empty_string=False
    if isinstance(typeTag, str):
        flag_empty_string = typeTag.strip() == '' 
    if flag_empty_string or typeTag==None:
        # if type not found in block then look for it in the neightbours
        
        # here the potential options sorted by their distances to block under analysis
        filtered_blocks,distances_with_block0 = findBlocksNearBlock(block,components )
        if len(filtered_blocks)>0:
            # select the nearest block
            i=-1
            for comp in filtered_blocks:
                i+=1
                typeTag_i = getTypeFromBlock(comp)
                if typeTag_i and isinstance(typeTag_i, str) and not(isTagBlock(comp)):
                    block_found = comp
                    return typeTag_i,distances_with_block0[i] , block_found
            
            # if not found type propose to use the name of the block
            i=-1
            for comp in filtered_blocks:
                i+=1
                typeTag_i = comp['block_name']
                if len(typeTag_i)>1 and isinstance(typeTag_i, str) and not(isTagBlock(comp)):
                    block_found = comp
                    return typeTag_i,distances_with_block0[i] , block_found
                
    return typeTag,distance, block_found


def parse_tag_code(tag_code):
    # Updated regex to capture an optional second alphabetic part at the end
    match = re.match(r"([A-Za-z]+)(\d+)([A-Za-z]*)$", tag_code)
    if match:
        target_object_type = match.group(1)  # The first alphabetic part (HV or FXS)
        target_object_loop_number = int(match.group(2))  # The numeric part (201)
        target_object_type_2nd = match.group(3)  # The optional second alphabetic part (X)
        
        return target_object_type, target_object_loop_number, target_object_type_2nd
    else:
        #raise ValueError(f"Invalid tag code format: {tag_code}")
        print(f"Warning: Invalid tag code format: {tag_code}")
        target_object_type = tag_code
        target_object_loop_number = None
        target_object_type_2nd = None
        return target_object_type, target_object_loop_number, target_object_type_2nd
        
def getTagCode(block):
    targetObjectType = ''
    targetObjectLoopNumber = ''
    
    
    if isTagBlock(block):
        targetObjectExtracted={'targetObjectType':None,
                               'targetObjectLoopNumber':None,
                               'targetObjectType2nd':None,
                               }
        attributes = block['attributes']
        
        # Define the search keys
        key_TargetObjectType = '#(TargetObject.Type)'
        key_TargetObjectLoopNumber = '#(TargetObject.LoopNumber)'
        key_TargetObjectTag = '#(TargetObject.Tag)'
        
        # Create a case-insensitive dictionary for attribute keys
        attributes_lower = {key.lower(): value for key, value in attributes.items()}
        
        # Convert search keys to lowercase
        key_TargetObjectType_lower = key_TargetObjectType.lower()
        key_TargetObjectLoopNumber_lower = key_TargetObjectLoopNumber.lower()
        key_TargetObjectTag_lower = key_TargetObjectTag.lower()
        
        # Attempt to fetch both values in a single step, avoiding redundant lookups
        targetObjectType = attributes_lower.get(key_TargetObjectType_lower,'')
        targetObjectLoopNumber = attributes_lower.get(key_TargetObjectLoopNumber_lower,'')
        targetObjectType2nd = ''
        
        tag_complete = targetObjectType + targetObjectLoopNumber
        targetObjectType, targetObjectLoopNumber, targetObjectType2nd = parse_tag_code(tag_complete)
        
        if targetObjectType=='': targetObjectType = None
        if targetObjectLoopNumber=='': targetObjectLoopNumber = None
        
        if targetObjectType or targetObjectLoopNumber:
            targetObjectExtracted.update({
                'targetObjectType':targetObjectType,
                'targetObjectLoopNumber':targetObjectLoopNumber,
                'targetObjectType2nd':targetObjectType2nd,
                })
            # Both were found, return them
            return targetObjectExtracted
        else:
            # Try to find the tag instead
            targetObjectTag = attributes_lower.get(key_TargetObjectTag_lower)
            if targetObjectTag:
                # Parse the tag to get type and loop number
                # if 'G' == targetObjectTag:
                #     print(block)
                targetObjectType, targetObjectLoopNumber, targetObjectType2nd = parse_tag_code(targetObjectTag)
                targetObjectExtracted.update({
                    'targetObjectType':targetObjectType,
                    'targetObjectLoopNumber':targetObjectLoopNumber,
                    'targetObjectType2nd':targetObjectType2nd,
                    })
                return targetObjectExtracted
            else:
                print(f"Warning: The provided block '{block['block_name']}' is missing TagObjects.")
                return None
                #raise ValueError(f"The provided block '{block['block_name']}' is missing TagObjects.")
    
    else:
        raise ValueError("The provided block is not a TagBlock.")


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
    text_val = text if text else ''

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
    text_val = text if text else ''
    
    # Check for an existing attribute with the specified tag.
    attrib = get_attrib_from_tag(tag, insert_block)
    
    if attrib:
        # If the attribute exists, update its text value.
        attrib.dxf.text = text_val
        return attrib
    else:
        # If the attribute does not exist, create a new one and synchronize with the block definition.
        new_attrib = add_new_tag_to_insert(tag, text=text_val, insert_block=insert_block)
        new_attrib.dxf.text = text_val
        return new_attrib