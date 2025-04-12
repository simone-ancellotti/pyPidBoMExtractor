# -*- coding: utf-8 -*-
"""
Created on Fri Apr 11 16:16:42 2025

@author: user
"""
from pyPidBoMExtractor.bom_generator import export_bom_to_excel, extract_bom_from_dxf,filterBOM_Ignore
from pyPidBoMExtractor.bom_generator import compare_bomsJSON,convert_bom_dxf_to_JSON,load_bom_from_excel_to_JSON
from pyPidBoMExtractor.utils import update_tag_value_in_block
import os
import json
import ezdxf

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
           
           

rows_xls_no = import_BOMjson_into_DXF(bom_revisedJSON,bom_dxf)        
        


# comp1['P&ID TAG']
# row=get_row_by_field(bom_revisedJSON, 'P&ID TAG', comp1['P&ID TAG'], case_sensitive=False)
# for key in tags_xls2dxf.keys():
#     key_dxf = tags_xls2dxf[key]
#     text = row.get(key)
#     update_tag_value_in_block(text, key_dxf, target_entity)

doc.saveas(r"./tests/test1/Schema di funzionamento_rev1.1_modified.dxf")
print("Modified DXF saved as 'my_drawing_modified.dxf'.")