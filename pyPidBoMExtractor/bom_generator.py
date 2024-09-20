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

    # Start writing BOM data after the header row (e.g., starting at row 2)
    start_row = 2
    
    for i, bom_item in enumerate(bom_data, start=start_row):
        # Assuming bom_item is a dictionary or list where you know the order of columns
        sheet.cell(row=i, column=1).value = bom_item['component_name']  # Example column 1
        sheet.cell(row=i, column=2).value = bom_item['quantity']        # Example column 2
        sheet.cell(row=i, column=3).value = bom_item['description']     # Example column 3

        # Add more columns based on the template structure

    # Save the workbook with a new name
    workbook.save(output_path)
    print(f"BOM successfully exported to {output_path}")
