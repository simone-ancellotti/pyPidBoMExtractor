import numpy as np
from ezdxf.math import BoundingBox, Vec3
import re

listTagBlockNames = ('Tag_block','Ball_Tag',
                     'STICKER Moving Machine', 'STICKER Equipment Name',
                     'Tag_Instrument')




def normalize(v):
    v=np.array(v)
    norm = np.linalg.norm(v)
    if norm == 0: return v
    return v / norm

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


def findBlocksNearBlock(block,components ):
    filtered_blocks = []
    distances_with_block0 = []
    positionBlock0 = np.array(block['insert_point'])
    mainDimBlock0  = getCaracteristicDimensionBlock(block)
    i=-1
    for component in components:
        i+=1
        if block['block_def'] != component['block_def']:
            positionComponent1 = np.array(component['insert_point'])
            mainDim1 = getCaracteristicDimensionBlock(component)
            
            criticalDistance = getCriticalDistance(mainDimBlock0,mainDim1)
            distance = np.linalg.norm(positionBlock0 - positionComponent1)
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
    else: 
        typeTag = attributes.get('DESCRIPTION')
    return typeTag

def findTypeBlockFromTag(block,components):
    #blocksValveAndMachine = [comp for comp in components ]
    # attributes = block['attributes']
    
    # typeTag = ''
    # if 'TYPE' in attributes.items():
    #     typeTag = attributes['TYPE']
    typeTag = getTypeFromBlock(block)
    distance = np.nan
    
    if typeTag == '' or typeTag==None:
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
                    return typeTag_i,distances_with_block0[i]
            
            # if not found type propose to use the name of the block
            i=-1
            for comp in filtered_blocks:
                i+=1
                typeTag_i = comp['block_name']
                if len(typeTag_i)>1 and isinstance(typeTag_i, str) and not(isTagBlock(comp)):
                    return typeTag_i,distances_with_block0[i]
                
    return typeTag,distance


def parse_tag_code(tag_code):
    # Updated regex to capture an optional second alphabetic part at the end
    match = re.match(r"([A-Za-z]+)(\d+)([A-Za-z]*)$", tag_code)
    if match:
        target_object_type = match.group(1)  # The first alphabetic part (HV or FXS)
        target_object_loop_number = int(match.group(2))  # The numeric part (201)
        target_object_type_2nd = match.group(3)  # The optional second alphabetic part (X)
        
        return target_object_type, target_object_loop_number, target_object_type_2nd
    else:
        raise ValueError(f"Invalid tag code format: {tag_code}")
        
def getTagCode(block):
    targetObjectType = ''
    targetObjectLoopNumber = ''
    
    
    if isTagBlock(block):
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
        targetObjectType = attributes_lower.get(key_TargetObjectType_lower)
        targetObjectLoopNumber = attributes_lower.get(key_TargetObjectLoopNumber_lower)
        targetObjectType2nd = ''
        if targetObjectType and targetObjectLoopNumber:
            # Both were found, return them
            return targetObjectType, targetObjectLoopNumber,targetObjectType2nd
        else:
            # Try to find the tag instead
            targetObjectTag = attributes_lower.get(key_TargetObjectTag_lower)
            if targetObjectTag:
                # Parse the tag to get type and loop number
                return parse_tag_code(targetObjectTag)
            else:
                raise ValueError(f"The provided block '{block['block_name']}' is missing TagObjects.")
    
    else:
        raise ValueError("The provided block is not a TagBlock.")

