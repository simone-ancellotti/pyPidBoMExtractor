import tkinter as tk
from tkinter import ttk

# Sample data (replace with your actual data)
sample_bom_data = {
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

class FilterableTable(ttk.Frame):
    def __init__(self, master, data, default_filter_column="P&ID TAG", **kwargs):
        """
        A Frame that contains widgets to filter and display data in a Treeview.
        
        Args:
            master: The parent widget.
            data (dict): Data dictionary where keys are row identifiers and values are dictionaries.
            default_filter_column (str): The default column to filter on.
        """
        super().__init__(master, **kwargs)
        self.data = data
        self.filtered_data = data.copy()

        # Determine columns from a sample row (assuming all rows have the same keys)
        sample_row = next(iter(data.values()))
        self.columns = list(sample_row.keys())

        # Create a filter frame for choosing column & typing filter text.
        filter_frame = ttk.Frame(self)
        filter_frame.pack(side="top", fill="x", padx=5, pady=2)

        # Combobox to choose filter column:
        ttk.Label(filter_frame, text="Filter Column:").pack(side="left", padx=5)
        self.filter_col_var = tk.StringVar()
        # Set default column:
        self.filter_col_var.set(default_filter_column)
        self.column_combobox = ttk.Combobox(
            filter_frame, textvariable=self.filter_col_var,
            values=self.columns, state="readonly", width=15
        )
        self.column_combobox.pack(side="left", padx=5)
        self.column_combobox.bind("<<ComboboxSelected>>", self.on_filter_change)

        # Entry to type filter text:
        ttk.Label(filter_frame, text="Filter:").pack(side="left", padx=5)
        self.filter_var = tk.StringVar()
        filter_entry = ttk.Entry(filter_frame, textvariable=self.filter_var)
        filter_entry.pack(side="left", fill="x", expand=True, padx=5)
        filter_entry.bind("<KeyRelease>", self.on_filter_change)

        # Frame for the Treeview and scrollbars.
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Create the Treeview widget
        self.tree = ttk.Treeview(table_frame, columns=self.columns, show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")
        
        # Configure scrollbars.
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=hsb.set)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Initial population of the Treeview.
        self.populate_tree()

    def populate_tree(self):
        # Clear existing rows.
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insert filtered data.
        for key, row in self.filtered_data.items():
            tags = ()
            # Example: if the "P&ID TAG" starts with "pH6", highlight the row.
            # This is just an example; modify your logic as needed.
            if str(row.get("P&ID TAG", "")).startswith("pH6"):
                tags = ('highlight',)
            values = tuple(row.get(col, "") for col in self.columns)
            self.tree.insert("", "end", values=values, tags=tags)
        self.tree.tag_configure('highlight', background='lightblue')

    def on_filter_change(self, event=None):
        # Retrieve the current filter text and selected column.
        filter_text = self.filter_var.get().lower()
        filter_column = self.filter_col_var.get()
        self.filtered_data = {}
        for key, row in self.data.items():
            value = str(row.get(filter_column, "")).lower()
            if filter_text in value:
                self.filtered_data[key] = row
        self.populate_tree()

# Example usage: integrated into a Notebook tab.

def create_ui_with_dynamic_filter():
    root = tk.Tk()
    root.title("Filtered Table with Dynamic Column Selection")
    root.geometry("800x600")

    # Create Notebook for tabs.
    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    # Tab 1: Main (for demonstration)
    tab1 = ttk.Frame(notebook)
    notebook.add(tab1, text="Main")
    ttk.Label(tab1, text="Main UI goes here", font=("Arial", 16)).pack(padx=10, pady=10)

    # Tab 2: Filterable table
    tab2 = ttk.Frame(notebook)
    notebook.add(tab2, text="BOM Table")
    filterable_table = FilterableTable(tab2, data=sample_bom_data, default_filter_column="P&ID TAG")
    filterable_table.pack(fill="both", expand=True)

    root.mainloop()

if __name__ == '__main__':
    create_ui_with_dynamic_filter()
