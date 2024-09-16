# -*- coding: utf-8 -*-
"""
Created on Thu Aug 22 15:42:22 2024

@author: user
"""

import ezdxf

def is_text_in_circle(circle, text_entity):
    """
    Check if a text entity is within a circle.
    """
    circle_center = circle.dxf.center
    circle_radius = circle.dxf.radius
    text_position = text_entity.dxf.insert
    
    # Calculate the distance from the text to the circle's center
    distance = ((text_position.x - circle_center.x) ** 2 + (text_position.y - circle_center.y) ** 2) ** 0.5
    
    # If the distance is less than the circle's radius, the text is inside the circle
    return distance < circle_radius

def extract_components_from_dwg(filename):
    # Load the DWG file
    dwg = ezdxf.readfile(filename)
    msp = dwg.modelspace()

    components = []

    # Find all circles and texts in the drawing
    circles = [entity for entity in msp if entity.dxftype() == 'CIRCLE']
    texts = [entity for entity in msp if entity.dxftype() in ['TEXT', 'MTEXT']]

    for circle in circles:
        for text in texts:
            if is_text_in_circle(circle, text):
                component = {
                    "circle_center": (circle.dxf.center.x, circle.dxf.center.y),
                    "circle_radius": circle.dxf.radius,
                    "text": text.text
                }
                components.append(component)
                break  # Assuming one text per circle

    return components

def generate_bom(components):
    bom = {}
    for component in components:
        number = component['text']
        if number in bom:
            bom[number] += 1
        else:
            bom[number] = 1
    return bom

if __name__ == "__main__":
    # Replace with the path to your DWG file
    dwg_filename = 'P&ID_simple.dxf'
    
    components = extract_components_from_dwg(dwg_filename)
    bom = generate_bom(components)
    
    # Print the BOM
    for component, count in bom.items():
        print(f"Component {component}: {count} instances")
