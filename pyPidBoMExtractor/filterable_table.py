# -*- coding: utf-8 -*-
"""
Created on Sun Apr 13 16:16:53 2025

@author: user
"""

import tkinter as tk
from tkinter import ttk
from .bom_generator import get_keys_from_pid_tag

class FilterableTable(ttk.Frame):
    def __init__(self, master, data, columns, mapping=None, filter_column_default=None, 
                 colour_mapping=None, column_widths=None, default_width=100,**kwargs):
        """
        A general-purpose widget that displays tabular data with filtering controls.
        
        Args:
            master: Parent widget.
            data (dict): Dictionary of data rows (keys can be row IDs, values are dictionaries).
            columns (list): List of column display names.
            mapping (dict, optional): Mapping from display column names to the data's keys.
                                      If omitted, it is assumed display names match the data keys.
            filter_column_default (str, optional): Default filter column (display name).
            column_widths (dict, optional): Dictionary mapping display column names to desired widths.
            default_width (int): Width to use if a column is not specified in column_widths.
        """
        super().__init__(master, **kwargs)
        self.data = data
        self.columns = columns
        
        # If no mapping is provided, assume the display names are identical to the keys in data.
        self.mapping = mapping if mapping is not None else {col: col for col in columns}
        self.mapping_reversed = {value: key for key, value in self.mapping.items()}
        self.colour_mapping = colour_mapping
        
        self.column_widths = column_widths if column_widths is not None else {}
        self.default_width = default_width
        
        # Set default filter column.
        self.filter_column = filter_column_default if filter_column_default else columns[0]
        
        # Configure custom selection color for Treeview
        style = ttk.Style(self)
        style.map("Treeview",
                  background=[("selected", "#e0f7fa")],   # Light cyan
                  foreground=[("selected", "black")])
        
        
        # Create filter controls.
        self._create_filter_controls()
        # Create the table (Treeview) and scrollbars.
        self._create_table()
        # Initially populate with all data.
        self.filtered_data = self.data.copy() if self.data else {}
        self._populate_table()

    def _create_filter_controls(self):
        filter_frame = ttk.Frame(self)
        filter_frame.pack(fill="x", padx=5, pady=2)
        
        ttk.Label(filter_frame, text="Filter Column:").pack(side="left", padx=5)
        self.filter_col_var = tk.StringVar(value=self.filter_column)
        self.column_combobox = ttk.Combobox(
            filter_frame, textvariable=self.filter_col_var,
            values=self.columns, state="readonly", width=15
        )
        self.column_combobox.pack(side="left", padx=5)
        self.column_combobox.bind("<<ComboboxSelected>>", self._on_filter_change)
        
        ttk.Label(filter_frame, text="Filter:").pack(side="left", padx=5)
        self.filter_text_var = tk.StringVar()
        filter_entry = ttk.Entry(filter_frame, textvariable=self.filter_text_var)
        filter_entry.pack(side="left", fill="x", expand=True, padx=5)
        filter_entry.bind("<KeyRelease>", self._on_filter_change)

    def _create_table(self):
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Create Treeview with provided display columns.
        self.tree = ttk.Treeview(table_frame, columns=self.columns, show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")
        for col in self.columns:
            self.tree.heading(col, text=col)
            width = self.column_widths.get(col, self.default_width)
            self.tree.column(col, width=width, anchor="center")
        
        self.tree.bind("<Control-c>", self._copy_selection_to_clipboard)
        self.tree.bind("<Double-1>", self._on_double_click)
        
        # Vertical scrollbar.
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vsb.set)
        # Horizontal scrollbar.
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=hsb.set)
        
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    def _on_filter_change(self, event=None):
        """
        Called when either the selected filter column or filter text changes.
        Updates the filtered data set.
        """
        selected_col = self.filter_col_var.get()  # Display name
        filter_key = self.mapping.get(selected_col, selected_col)
        filter_text = self.filter_text_var.get().lower()
        
        self.filtered_data = {}
        for row_id, row in self.data.items():
            cell_value = str(row.get(filter_key, "")).lower()
            if filter_text in cell_value:
                self.filtered_data[row_id] = row
        self._populate_table()
    
    def _populate_table(self):
        """
        Clears and repopulates the Treeview based on filtered_data.
        """
        # Remove current data
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        
        # Insert filtered rows
        for row_id, row in self.filtered_data.items():
            # Create a tuple of values in the order of self.mapping (using display columns)
            values = tuple(row.get(self.mapping[col], "") for col in self.columns)
            tags = ()
            # Example: add a tag to highlight rows based on a specific condition.
            # Here we check for the column "P&ID TAG" if present.
            pid_key = self.mapping.get("P&ID TAG", "P&ID TAG")
            # if str(row.get(pid_key, "")).startswith("XV"):
            #     tags = ("highlight",)
            value_pid= str(row.get(pid_key, ""))
            if self.colour_mapping:
                color = self.colour_mapping.get(value_pid)
                if color:
                    tags = (color,)
                    
            self.tree.insert("", "end",iid=str(row_id), values=values, tags=tags)
        #self.tree.tag_configure("highlight", background="lightblue")
        if self.colour_mapping:
            for c in self.colour_mapping.values():
                self.tree.tag_configure(c, background=c)
                
    def _copy_selection_to_clipboard(self, event=None):
        selection = self.tree.selection()
        if not selection:
            return
    
        rows = []
        for item_id in selection:
            row_values = self.tree.item(item_id)["values"]
            row_str = "\t".join(str(v) for v in row_values)
            rows.append(row_str)
    
        result = "\n".join(rows)
        self.clipboard_clear()
        self.clipboard_append(result)
        self.update()  # Ensures clipboard is updated

    
    def _on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
    
        row_iid = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        col_index = int(col_id.replace("#", "")) - 1
    
        x, y, width, height = self.tree.bbox(row_iid, col_id)
        current_value = self.tree.set(row_iid, column=col_id)
    
        entry = tk.Entry(self.tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_value)
        entry.focus()
    
        def on_confirm(event=None):
            new_value = entry.get()
            self.tree.set(row_iid, column=col_id, value=new_value)
    
            try:
                data_row_id = int(row_iid)  # the key in self.data
                column_name = self.columns[col_index]  # display column
                data_key = self.mapping.get(column_name, column_name)  # actual dict key
                #print(f"Updating row {data_row_id}, key '{data_key}' with '{new_value}'")
                if data_row_id in self.data and (current_value!=new_value):
                    self.data[data_row_id][data_key] = new_value  # 🔥 this is the real update
                    if data_key=='P&ID TAG':
                        pid_key_tag = new_value
                        tag_comps = get_keys_from_pid_tag(pid_key_tag)
                        if tag_comps[0] != '@':
                            L, N, D = tag_comps
                            data_key_L = self.mapping.get('L', 'L')  # actual dict key
                            data_key_N = self.mapping.get('N','N')  # actual dict key
                            data_key_D = self.mapping.get('D','D')  # actual dict key
                            self.data[data_row_id][data_key_L] = str(L)
                            self.data[data_row_id][data_key_N] = str(N)
                            self.data[data_row_id][data_key_D] = str(D)
                            self._populate_table()
                    
                    print(f"Updating row {data_row_id}, key '{data_key}' with '{new_value}'")

            except Exception as e:
                print(f"Update failed: {e}")
    
            entry.destroy()
    
        entry.bind("<Return>", on_confirm)
        entry.bind("<FocusOut>", lambda e: entry.destroy())
        column_name = self.columns[col_index]  # display column
        data_key = self.mapping.get(column_name, column_name)  # actual dict key
        print(f"check value updated {self.data[int(row_iid)][data_key]}")


    def set_data(self, new_data,new_colour_mapping=None):
        """
        Updates the table with new data.
        
        Args:
            new_data (dict): The new data dictionary.
        """
        self.data = new_data
        self.colour_mapping = new_colour_mapping
        self.filtered_data = self.data.copy() if self.data else {}
        self._populate_table()
