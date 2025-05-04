# -*- coding: utf-8 -*-
"""
Created on Fri May  2 10:25:45 2025

@author: user
"""
import win32com.client
import pythoncom

def zoom_to_point_fixed(x, y, zoom_factor=50.0):
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        #doc.ActiveSpace = 1  # Ensure we're in ModelSpace
        doc.SendCommand("_UCS\nW\n")  # Force WCS

        vp = doc.ActiveViewport

        print("Before Zoom:", vp.Center, vp.Height, vp.Width)

        center_point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y)])
        vp.Center = center_point
        vp.Height = zoom_factor
        vp.Width = zoom_factor
        doc.ActiveViewport = vp
        doc.Regen(1)

        print("✅ Zoomed to:", center_point)
        print("After Zoom:", vp.Center, vp.Height, vp.Width)

    except Exception as e:
        print("❌ AutoCAD zoom failed:", e)

        #Ýprint(f"✅ Zoomed to ({x}, {y}) with zoom {zoom_factor}")


def zoom_to_point_command(x, y, height=50):
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



# Example usage
#zoom_to_point_fixed(661, 923)
zoom_to_point_command(661, 923)
