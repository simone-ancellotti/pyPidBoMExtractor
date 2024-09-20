from .utils import isTagBlock, getTagCode, findTypeBlockFromTag
from .extractor import extract_blocks_with_attributes_and_dimensions
import openpyxl
import pandas as pd

def generate_bom(components):
    
    #tagBlocks = [c for c in components if isTagBlock(c)]
    bom = {}
    i=-1
    for component in components:
        i+=1
        number = i+1
        #block_name = component['block_name']
        attributes = component['attributes']
        #contains_circle = component['contains_circle']
        
        targetObjectType, targetObjectLoopNumber, targetObjectType2nd = getTagCode(component)     
        typeTag,distance = findTypeBlockFromTag(component,components)
        description = attributes.get('DESCRIPTION')
        
        bom.update( {number:{
                "count":number,
                'targetObjectType':targetObjectType,
                'targetObjectLoopNumber':targetObjectLoopNumber,
                'targetObjectType2nd':targetObjectType2nd,
                'TYPE':typeTag,
                'description':description
                }})
    return bom

def print_bom(bom):
    print("\nGenerated BOM:\n")
    print(f"{'#':<3}|{'L':<4}|{'N':<5}|{'D':<4}|{'P&ID TAG':<10}|{'Type':<30}|{'Description':<30}")
    print("-" * 40)
    for it in bom.keys():
        component = bom[it]
        L = component['targetObjectType']
        N = component['targetObjectLoopNumber']
        D = component['targetObjectType2nd']
        pid_TAG = str(L)+str(N)+str(D)
        comp_type = component['TYPE']
        description=component['description']
        if description == None: description=''
        print(f"{it:<3}|{L:<4}|{N:<5}|{D:<4}|{pid_TAG:<10}|{comp_type:<30}|{description:<30}")



def export_bom_to_excel(bom_data, template_path, output_path):
    # Load the template file
    workbook = openpyxl.load_workbook(template_path)
    sheet = workbook.active  # Assuming the BOM sheet is the active one
    
    # Mapping between BOM keys and Excel column headers
    header_mapping = {
        'count': '#',
        'targetObjectType': 'L',
        'targetObjectLoopNumber': 'N',
        'targetObjectType2nd': 'D',
        'TYPE': 'Type',
        'description': 'Description'
    }
    
    # Find the corresponding column index for each header in the Excel template
    header_row = 1  # Assuming headers are in the first row
    column_indexes = {}
    
    for col in range(1, sheet.max_column + 1):
        header_value = sheet.cell(row=header_row, column=col).value
        for bom_key, header_name in header_mapping.items():
            if header_value == header_name:
                column_indexes[bom_key] = col
        if header_value == 'P&ID TAG':
            pid_tag_col = col  # Find the column for "P&ID TAG"
    
    # Start writing BOM data after the header row (e.g., starting at row 2)
    start_row = 2
    for i, bom_item in enumerate(bom_data.values(), start=start_row):
        # Write the regular BOM fields
        for bom_key, value in bom_item.items():
            if bom_key in column_indexes:
                col = column_indexes[bom_key]
                sheet.cell(row=i, column=col).value = value
        
        # Generate and write the "P&ID TAG" (L + N + D)
        l_value = bom_item.get('targetObjectType', '')
        n_value = str(bom_item.get('targetObjectLoopNumber', ''))
        d_value = bom_item.get('targetObjectType2nd', '')
        
        pid_tag = f"{l_value}{n_value}{d_value}"
        sheet.cell(row=i, column=pid_tag_col).value = pid_tag
    
    # Save the workbook with a new name
    workbook.save(output_path)
    print(f"BOM successfully exported to {output_path}")

def extract_bom_from_dxf(dwg_file):
    components = extract_blocks_with_attributes_and_dimensions(dwg_file)
    
    # filter only the blocks stickers TAGs to determine the components in BOM
    # feel free to change the logic
    tagBlocks = [c for c in components if isTagBlock(c)]
    bom = generate_bom(tagBlocks)
    # Print the formatted BOM with dimensions
    print_bom(bom)
    return bom
    

def load_bom_from_excel(file_path):
    """ Load BOM data from Excel file. """
    df = pd.read_excel(file_path)
    
    # Assuming columns like 'L', 'N', 'D', 'Type', 'Description'
    df['N'] = df['N'].fillna(0).astype(int).astype(str)  # Convert N to int, then to string
    df['D'] = df['D'].fillna('').astype(str)  # Convert D to empty string if NaN
    
    # Construct P&ID TAG by concatenating L, N, and D
    df['P&ID TAG'] = df['L'].astype(str) + df['N'] + df['D']
    
    return df

def convert_bom_dxf_to_dataframe(bom_dxf):
    """ Convert BOM dictionary from DXF to a pandas DataFrame. """
    if not isinstance(bom_dxf, dict):
        raise ValueError("Input must be a dictionary.")
        
    # Create a list of dictionaries for each item in the BOM
    bom_list = []
    for item in bom_dxf.values():
        bom_list.append({
            'count': item['count'],
            'targetObjectType': item['targetObjectType'],
            'targetObjectLoopNumber': item['targetObjectLoopNumber'],
            'targetObjectType2nd': item['targetObjectType2nd'],
            'Type': item['TYPE'],
            'Description': item['description']
        })
    
    # Convert the list to a DataFrame
    bom_df = pd.DataFrame(bom_list)

    # Create the P&ID TAG column
    bom_df['P&ID TAG'] = (
        bom_df['targetObjectType'].astype(str) +
        bom_df['targetObjectLoopNumber'].astype(str) +
        bom_df['targetObjectType2nd'].astype(str)
    )

    return bom_df



def compare_boms(bom_dxf_df, bom_revised):
    """ Compare two BOMs and print differences. """
    
    # Convert bom_dxf to a DataFrame

    # Filter out rows where L, N, D, or P&ID TAG are empty
    # bom_dxf_df = bom_dxf_df.dropna(subset=['targetObjectType', 'targetObjectLoopNumber', 'targetObjectType2nd', 'P&ID TAG'])
    # bom_dxf_df = bom_dxf_df[bom_dxf_df[['targetObjectType', 'targetObjectLoopNumber', 'targetObjectType2nd', 'P&ID TAG']].applymap(lambda x: x not in [None, ''])]

    # Filter the revised BOM similarly
    bom_revised = bom_revised.dropna(subset=['L', 'N', 'D', 'P&ID TAG'])
    
    # Create sets of P&ID TAGs for comparison
    bom_dxf_set = set(bom_dxf_df['P&ID TAG'])
    bom_revised_set = set(bom_revised['P&ID TAG'])

    # Components in BOM_dxf but not in BOM_revised
    missing_in_revised = bom_dxf_set - bom_revised_set
    if missing_in_revised:
        print("\nComponents in DXF but not in the revised BOM:")
        for item in missing_in_revised:
            print(item)

    # Components in BOM_revised but not in BOM_dxf
    missing_in_dxf = bom_revised_set - bom_dxf_set
    if missing_in_dxf:
        print("\nComponents in revised BOM but not in DXF BOM:")
        for item in missing_in_dxf:
            print(item)

    # Optional: Compare other attributes (Type, Description, etc.)
    print("\nComparing attributes of matching components...")
    for tag in bom_dxf_set.intersection(bom_revised_set):
        dxf_row = bom_dxf_df[bom_dxf_df['P&ID TAG'] == tag].iloc[0]
        revised_row = bom_revised[bom_revised['P&ID TAG'] == tag].iloc[0]
        
        # Check for differences in attributes (Type, Description, etc.)
        for column in ['Type', 'Description']:  # Add more columns if needed
            dxf_value = dxf_row[column] if pd.notna(dxf_row[column]) else None
            if dxf_value == '': dxf_value = None
            revised_value = revised_row[column] if pd.notna(revised_row[column]) else None
            
            if dxf_value != revised_value:
                print(f"Component {tag}: {column} differs. DXF: {dxf_value}, Revised: {revised_value}")
