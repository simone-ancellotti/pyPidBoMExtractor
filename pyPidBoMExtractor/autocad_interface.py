# -*- coding: utf-8 -*-
"""
Created on Fri May  2 11:04:08 2025

@author: user
"""

import win32com.client
import pythoncom

def zoom_in_autocad(x, y, zoom_factor=30):
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        vp = doc.ActiveViewport

        center = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y)]
        )
        vp.Center = center
        vp.Height = zoom_factor * 2
        vp.Width = zoom_factor * 2
        doc.ActiveViewport = vp
        doc.Regen(1)
        print(f"Zoomed to ({x}, {y}) in AutoCAD")
    except Exception as e:
        print("AutoCAD zoom failed:", e)

def zoom_in_autocad_command(x, y, height=50):
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        #doc.ActiveSpace = 0  # Ensure we're in ModelSpace

        # Send the ZOOM -CENTER command
        command_str = f'_ZOOM C {x},{y} {height}\n'
        doc.SendCommand(command_str)

        print(f"✅ Sent zoom command to center on ({x}, {y}) with height {height}")
    except Exception as e:
        print("❌ AutoCAD zoom failed:", e)