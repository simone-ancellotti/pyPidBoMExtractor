# -*- coding: utf-8 -*-
"""
Created on Fri Apr 11 16:16:42 2025

@author: user
"""
from pyPidBoMExtractor.bom_generator import export_bom_to_excel, extract_bom_from_dxf,filterBOM_Ignore,sort_rows_by_pid_tag
from pyPidBoMExtractor.bom_generator import compare_bomsJSON,convert_bom_dxf_to_JSON,load_bom_from_excel_to_JSON,sort_bom_by_pid_tag
from pyPidBoMExtractor.utils import update_tag_value_in_block
import os
import json
import ezdxf
import re

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

dwg_file = r"./tests/test1/Schema di funzionamento_rev1.1.dxf"
bom_dxf,doc = extract_bom_from_dxf(dwg_file)

file_xls_path = r"./tests/test1/BOM_1.1_rev.xlsx"
bom_revisedJSON = load_bom_from_excel_to_JSON(file_xls_path)




comp1 = bom_dxf[1]
comp2 = bom_dxf[2]
target_entity = comp1['target_entity']
sticker_entity1 = comp1['tag_entity']['entity']

msp = doc.modelspace()
for att in sticker_entity1.attribs:
    print(att.dxfattribs())
[att.dxfattribs()['tag'] for att in target_entity.attribs]

tags_xls2dxf= {'Type':'TYPE', 'C type':'CONNECTIONTYPE', 'Description':'DESCRIPTION',
               'Fluid':'FLUID', 'Unit':'UNIT', 'Skid':'SKID', 'Type':'TYPE', 
               'Material':'MATERIAL', 'Seal Mat.':'SEAL_MAT', 
               'P     (kW)':'POWER_KW','PN   (bar)':'PN_bar', 
               'Act NO/NC':'ACT_NO_NC', 'Size':'SIZE','cap. (tanks L)':'CAP.(TANK L)', 
               'Q (m3/h)':'Q(m3/h)', 'Supplier':'SUPPLIER', 'Brand':'BRAND',
               'Model':'MODEL', 'Notes':'NOTES', 'Datasheet':'DATASHEET'
               }
tags_dxf2xls = {value: key for key, value in tags_xls2dxf.items()}




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
           
           

#rows_xls_no = import_BOMjson_into_DXF(bom_revisedJSON,bom_dxf)        
        


# comp1['P&ID TAG']
# row=get_row_by_field(bom_revisedJSON, 'P&ID TAG', comp1['P&ID TAG'], case_sensitive=False)
# for key in tags_xls2dxf.keys():
#     key_dxf = tags_xls2dxf[key]
#     text = row.get(key)
#     update_tag_value_in_block(text, key_dxf, target_entity)

doc.saveas(r"./tests/test1/Schema di funzionamento_rev1.1_modified.dxf")
print("Modified DXF saved as 'my_drawing_modified.dxf'.")

#

# def sort_key_for_pid_tag(tag):
#     """
#     Generate a sorting key from a P&ID TAG string.

#     The key is a tuple: (prefix, number, suffix) where:
#       - prefix is the initial alphabetic part in lowercase,
#       - number is the numeric part parsed as an integer,
#       - suffix is the remaining alphabetic part in lowercase.
      
#     If the tag doesn't match the expected pattern, it returns a key
#     that sorts the tag at the end.
#     """
#     pattern = re.compile(r"^([A-Za-z]+)(\d+)([A-Za-z]*)$")
#     match = pattern.match(tag or "")
#     if match:
#         prefix, num_str, suffix = match.groups()
#         return (prefix.lower(), int(num_str), suffix.lower())
#     else:
#         return (float('inf'), tag.lower())

# def sort_bom_by_pid_tag(bom_dict):
#     """
#     Sort the BOM dictionary by the 'P&ID TAG' field using sort_key_for_pid_tag.
    
#     Args:
#         bom_dict (dict): A dictionary where values are row dictionaries containing
#                          a "P&ID TAG" field.
                         
#     Returns:
#         dict: A new dictionary with entries sorted by the 'P&ID TAG'.
#               (Python 3.7+ dictionaries preserve insertion order.)
#     """
#     sorted_items = sorted(
#         bom_dict.items(),
#         key=lambda item: sort_key_for_pid_tag(item[1].get("P&ID TAG", ""))
#     )
#     return {key: value for key, value in sorted_items}

# # Example usage:


file_xls_path = r"G:/My Drive/ULIX tecnico/OFF/OFF 49 Kingspan NL/P&ID/OFF-49_P&ID_Kingspan_NL_rev1.0_bom.xlsx"
#file_xls_path = r"./tests/test1/BOM_1.1_rev.xlsx"
bom_revisedJSON2 = load_bom_from_excel_to_JSON(file_xls_path)
print([i['P&ID TAG'] for i in bom_revisedJSON2.values()])

def find_duplicates(bom_dict, field):
    """
    Find duplicate rows in bom_dict based on a given field.

    Args:
        bom_dict (dict): A dictionary where keys are row identifiers and values are dictionaries.
        field (str): The field name to check duplicates on (e.g., "P&ID TAG").

    Returns:
        dict: A dictionary mapping each duplicate field value to a list of row IDs (keys) that share that value.
              Only field values that occur more than once are included.
    """
    seen = {}
    duplicates = {}
    
    for row_id, row in bom_dict.items():
        value = row.get(field)
        if value in seen:
            # First duplicate occurrence: add the already-seen row.
            if value not in duplicates:
                duplicates[value] = [seen[value]]
            duplicates[value].append(row_id)
        else:
            seen[value] = row_id

    return duplicates

bom_dict={1:{'P&ID TAG':'XV100'},2:{'P&ID TAG':'XV100'},3:{'P&ID TAG':'XV110'}}
find_duplicates(bom_dict, 'P&ID TAG')