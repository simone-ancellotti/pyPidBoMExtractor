from .utils import isTagBlock, getTagCode, findTypeBlockFromTag
import openpyxl

def generate_bom(components):
    
    tagBlocks = [c for c in components if isTagBlock(c)]
    bom = {}
    i=-1
    for component in tagBlocks:
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
