# -*- coding: utf-8 -*-
"""
Created on Fri Apr 11 13:00:01 2025

@author: user
"""

import ezdxf
from .utils import update_tag_value_in_block

tags_xls2dxf= {'Type':'TYPE', 'C type':'CONNECTIONTYPE', 'Description':'DESCRIPTION',
               'Fluid':'FLUID', 'Unit':'UNIT', 'Skid':'SKID', 'Type':'TYPE', 
               'Material':'MATERIAL', 'Seal Mat.':'SEAL_MAT', 
               'P     (kW)':'POWER_KW','PN   (bar)':'PN_bar', 
               'Act NO/NC':'ACT_NO_NC', 'Size':'SIZE','cap. (tanks L)':'CAP.(TANK L)', 
               'Q (m3/h)':'Q(m3/h)', 'Supplier':'SUPPLIER', 'Brand':'BRAND',
               'Model':'MODEL', 'Notes':'NOTES', 'Datasheet':'DATASHEET'
               }
tags_dxf2xls = {value: key for key, value in tags_xls2dxf.items()}


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
                        key_dxf = tags_xls2dxf[key_xls]
                        text_xls = item_xls.get(key_xls)
                        text_xls = text_xls.strip()
                        update_tag_value_in_block(text_xls, key_dxf, entity_dxf_to_modify)
                else: 
                    rows_xls_no.append()
    return rows_xls_no
           