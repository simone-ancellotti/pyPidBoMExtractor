# -*- coding: utf-8 -*-
"""
Created on Fri Apr 11 13:00:01 2025

@author: user
"""

import ezdxf
from .utils import update_tag_value_in_block,getTagCode,parse_tag_code
from .bom_generator import tags_xls2dxf,tags_dxf2xls

key_TargetObjectType = '#(TargetObject.Type)'
key_TargetObjectLoopNumber = '#(TargetObject.LoopNumber)'
key_TargetObjectTag = '#(TargetObject.Tag)'

def get_row_by_field(bom_dict, field_name, field_value, case_sensitive=True):
    """
    Search for the first row in the BOM dictionary that matches the given field name and value.
    
    Args:
        bom_dict (dict): A dictionary representing the BOM (e.g., read from an Excel file),
                         where keys are row numbers and values are dictionaries of column values.
        field_name (str): The key (i.e., column name) to search in, e.g., 'P&ID TAG'.
        field_value: The value to look for within that column.
        case_sensitive (bool): If False, the comparison is done case-insensitively. Defaults to True.
    
    Returns:
        dict or None: The dictionary corresponding to the first row that matches; otherwise, None.
    """
    for row in bom_dict.values():
        current_value = row.get(field_name)
        if current_value is None:
            continue

        if case_sensitive:
            if current_value == field_value:
                return row
        else:
            # Convert both to strings and compare in lowercase.
            if str(current_value).strip().replace(' ','').lower() == str(field_value).strip().replace(' ','').lower():
                return row

    return None

def import_BOMjson_into_DXF(bom_revisedJSON,bom_dxf):
    rows_xls_no = []
    for item_xls in bom_revisedJSON.values():
        PID_TAG_xls = item_xls.get('P&ID TAG')
        if PID_TAG_xls:
            dxf_item_found = get_row_by_field(bom_dxf, 'P&ID TAG', PID_TAG_xls , case_sensitive=False)
            if dxf_item_found:
                entity_dxf_to_modify = dxf_item_found['target_entity']           
                if not(entity_dxf_to_modify):
                    entity_dxf_to_modify = dxf_item_found['tag_entity']['entity']
                
                if entity_dxf_to_modify:
                    for key_xls in tags_xls2dxf.keys():
                        if key_xls != 'P&ID TAG':
                            key_dxf = tags_xls2dxf[key_xls]
                            text_xls = str(item_xls.get(key_xls))
                            text_xls = text_xls.strip()
                            update_tag_value_in_block(text_xls, key_dxf, entity_dxf_to_modify)
                else: 
                    rows_xls_no.append()
    return rows_xls_no
           
def update_dxfJSON_into_dxf_drawing(bom_dxf):
    for row_id, dxf_item in bom_dxf.items():
        flagSynchronized = dxf_item.get("flagSynchronized",True)
        if not(flagSynchronized):
            tag_entity = dxf_item['tag_entity']['entity']
            entity_dxf_to_modify = dxf_item['target_entity']
            if not(entity_dxf_to_modify):
                entity_dxf_to_modify = tag_entity
                
            if entity_dxf_to_modify:
                for key_dxf in tags_dxf2xls.keys():
                    print(key_dxf)
                    if key_dxf != 'P&ID TAG':
                        text_dxf_updated = str(dxf_item.get(key_dxf))
                        text_dxf_updated = text_dxf_updated.strip()
                        update_tag_value_in_block(text_dxf_updated, key_dxf, entity_dxf_to_modify)
                    else:
                        pid_tag_value_new = str(dxf_item.get(key_dxf))
                        #print(getTagCode(dxf_item))
                        list_all_tags =[att.dxf.tag.upper() for att in tag_entity.attribs]
                        typeTag_entity = ''
                        if key_TargetObjectTag.upper() in list_all_tags:
                            typeTag_entity = 'sticker'
                        else:
                            required_keys = [key_TargetObjectLoopNumber, key_TargetObjectType]
                            if all(tag.upper() in list_all_tags for tag in required_keys):
                                typeTag_entity = 'ball'
                        L, N, D  = parse_tag_code(pid_tag_value_new)
                        N = str(N)
                        for att in tag_entity.attribs:
                            if typeTag_entity == 'sticker':
                                if att.dxf.tag.upper() == key_TargetObjectTag.upper():
                                    if att.dxf.text.upper() != pid_tag_value_new.upper():
                                        att.dxf.text = pid_tag_value_new
                            if typeTag_entity == 'ball':
                                if att.dxf.tag.upper() == key_TargetObjectType.upper():
                                    if att.dxf.text.upper() != L.upper():
                                        att.dxf.text = L
                                if att.dxf.tag.upper() == key_TargetObjectLoopNumber.upper():
                                    if str(att.dxf.text).upper() != N.upper():
                                        att.dxf.text = N
                                    

                        
                                
                        
            print(dxf_item)
        
    return None
    