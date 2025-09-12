# -*- coding: utf-8 -*-
"""
Created on Wed Apr 16 15:57:00 2025

@author: user
"""

#!/usr/bin/env python3
import os
import shutil
import sys

def copy_folder(src, dst, skip_list):
    """
    Copies the contents of the src directory to the dst directory,
    skipping any subfolders whose names are in skip_list.
    
    Args:
        src (str): The source directory path.
        dst (str): The destination directory path.
        skip_list (list): A list of subfolder names to skip.
                          (Only the name of the folder is matched,
                          not the full path.)
    """
    if not os.path.exists(src):
        raise Exception(f"Source directory '{src}' does not exist.")
    
    # Create the destination directory if it doesn't exist.
    if not os.path.exists(dst):
        os.makedirs(dst)
    
    for root, dirs, files in os.walk(src):
        # Compute the relative path from the source directory.
        rel_path = os.path.relpath(root, src)
        # Determine the destination directory for the current folder.
        dest_dir = os.path.join(dst, rel_path) if rel_path != '.' else dst
        
        # Remove directories that are in the skip list.
        # This modifies 'dirs' in-place, so os.walk will not traverse them.
        dirs[:] = [d for d in dirs if d not in skip_list]
        
        # Create the destination subfolder if needed.
        if not os.path.exists(dest_dir):
            os.makedirs(dest_dir)
        
        # Copy each file from the current directory.
        for file in files:
            src_file = os.path.join(root, file)
            dst_file = os.path.join(dest_dir, file)
            shutil.copy2(src_file, dst_file)
            print(f"Copied: {src_file} -> {dst_file}")

src = r'C:\ASETS-code\python\pyp-id_finder'
dst = r'G:\My Drive\ULIX tecnico\Varie\Python\pyp-id_finder'
skip_list=['.git','build']
copy_folder(src, dst, skip_list)

# for ext in ['dwg','dxf']:
#     template_dwg_file = os.path.join(src, r'templates\P&ID_simple_template.'+ext)
#     template_dwg_file_dest = os.path.join(dst, r'templates\P&ID_simple_template.'+ext)
#     shutil.copy2(template_dwg_file, template_dwg_file_dest)
