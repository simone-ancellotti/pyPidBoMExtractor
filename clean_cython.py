# -*- coding: utf-8 -*-
"""
Created on Tue Apr 29 10:49:01 2025

@author: user
"""

import os
import shutil

def clean_cython_generated_files(folder):
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith((".c", ".pyd", ".so")):
                filepath = os.path.join(root, file)
                print(f"Deleting: {filepath}")
                os.remove(filepath)
    # Also remove 'build' folder
    build_path = os.path.join(folder, "build")
    if os.path.exists(build_path):
        print(f"Deleting folder: {build_path}")
        shutil.rmtree(build_path)

if __name__ == "__main__":
    clean_cython_generated_files(".")  # "." means current folder
