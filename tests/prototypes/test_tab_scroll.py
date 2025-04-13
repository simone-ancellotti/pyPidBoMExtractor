# -*- coding: utf-8 -*-
"""
Created on Sun Apr 13 11:03:29 2025

@author: user
"""

import tkinter as tk
from tkinter import ttk

# Sample dictionary data.
bom_revisedJSON = {
    37: {'#': 30,
         'L': 'pH',
         'N': '601',
         'D': '',
         'P&ID TAG': 'pH601',
         'Fluid': '',
         'Unit': '',
         'DI': '',
         'DO': '',
         'AI': '',
         'AO': '',
         'Skid': '',
         'Type': 'pH metro',
         'Description': '',
         'Material': 'AISI345',
         'Seal Mat.': '',
         'P     (kW)': '',
         'PN   (bar)': '',
         'Act NO/NC': '',
         'Size': '',
         'C type': ' Flanged',
         'cap. (tanks L)': '',
         'Q (m3/h)': '',
         'Supplier': '',
         'Brand': '',
         'Model': '',
         'Notes': '',
         'OFFER REQUEST': '',
         'OFFER RECEIVED': '',
         'ORDER': '',
         'PRICE': '',
         'PAID': '',
         'ARRIVED': '',
         'Datasheet': '',
         'count': 37},
    38: {'#': 31,
         'L': 'pH',
         'N': '931',
         'D': '',
         'P&ID TAG': 'pH931',
         'Fluid': '',
         'Unit': '',
         'DI': '',
         'DO': '',
         'AI': '',
         'AO': '',
         'Skid': '',
         'Type': 'pH metro',
         'Description': '',
         'Material': 'AISI346',
         'Seal Mat.': '',
         'P     (kW)': '',
         'PN   (bar)': '',
         'Act NO/NC': '',
         'Size': '',
         'C type': ' Flanged',
         'cap. (tanks L)': '',
         'Q (m3/h)': '',
         'Supplier': '',
         'Brand': '',
         'Model': '',
         'Notes': '',
         'OFFER REQUEST': '',
         'OFFER RECEIVED': '',
         'ORDER': '',
         'PRICE': '',
         'PAID': '',
         'ARRIVED': '',
         'Datasheet': '',
         'count': 38}
}

def create_tabbed_ui():
    root = tk.Tk()
    root.title("Tabbed UI with Scrollable Table")
    root.geometry("800x600")

    # Create a Notebook widget for tabs.
    notebook = ttk.Notebook(root)
    notebook.pack(fill='both', expand=True)

    # -----------------------
    # Tab 1: Simple Info Tab
    # -----------------------
    tab1 = ttk.Frame(notebook)
    notebook.add(tab1, text="Tab 1")
    tk.Label(tab1, text="This is Tab 1", font=("Arial", 16)).pack(padx=10, pady=10)

    # ---------------------------
    # Tab 2: Table Display Tab
    # ---------------------------
    tab2 = ttk.Frame(notebook)
    notebook.add(tab2, text="Data Table")

    # Create a frame to hold the table and scrollbars.
    table_frame = ttk.Frame(tab2)
    table_frame.pack(fill='both', expand=True, padx=10, pady=10)

    # Create the Treeview widget for displaying table data.
    # Determine columns from the keys of one of the BOM rows.
    sample_row = bom_revisedJSON[next(iter(bom_revisedJSON))]
    columns = list(sample_row.keys())
    tree = ttk.Treeview(table_frame, columns=columns, show="headings")
    tree.grid(row=0, column=0, sticky="nsew")

    # Set up the column headings and basic column options.
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor="center")

    # Insert rows into the tree.
    for row in bom_revisedJSON.values():
        # Generate a tuple of values in the order of columns.
        values = tuple(row.get(col, "") for col in columns)
        tree.insert("", "end", values=values)

    # Add vertical scrollbar.
    vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    vsb.grid(row=0, column=1, sticky="ns")
    tree.configure(yscrollcommand=vsb.set)

    # Add horizontal scrollbar.
    hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
    hsb.grid(row=1, column=0, sticky="ew")
    tree.configure(xscrollcommand=hsb.set)

    # Configure grid weight for proper resizing.
    table_frame.rowconfigure(0, weight=1)
    table_frame.columnconfigure(0, weight=1)

    root.mainloop()

if __name__ == '__main__':
    create_tabbed_ui()
