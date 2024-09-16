# -*- coding: utf-8 -*-
"""
Created on Thu Aug 22 15:47:58 2024

@author: user
"""

import win32com.client
import os

def convert_dwg_to_dxf(dwg_file_path, output_dxf_path):
    # Check if the file exists
    if not os.path.isfile(dwg_file_path):
        raise FileNotFoundError(f"The file {dwg_file_path} does not exist.")
    
    # Create an AutoCAD application object
    acad = win32com.client.Dispatch("AutoCAD.Application")
    
    try:
        # Open the DWG file
        doc = acad.Documents.Open(dwg_file_path)
    
        # Save the file as DXF
        doc.SaveAs(output_dxf_path, 25)  # 16 represents the DXF R12 format
        
        # Close the document
        doc.Close()
    
    except Exception as e:
        print(f"An error occurred: {e}")
    
    finally:
        # Quit AutoCAD application
        acad.Quit()

if __name__ == "__main__":
    dwg_file = r'C:\ASETS-code\python\pyP&ID_finder\P&ID_ver9.dwg'
    dxf_output = r'C:\ASETS-code\python\pyP&ID_finder\P&ID_ver9.dxf'

    convert_dwg_to_dxf(dwg_file, dxf_output)
    print(f"Converted {dwg_file} to {dxf_output}")