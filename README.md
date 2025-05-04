## 🧰 How to Use

1. **Launch the App**
   Run this command to start the graphical user interface:

   ```bash
   python bom_extractor_ui.py
   ```

2. **Upload Files**

   * **Upload DXF**: Click "Upload DXF" and choose your P\&ID drawing.
   * **Upload Template Excel**: Click "Upload Template Excel" and select your BOM template.
   * **Upload Revised BOM Excel**: Optional, for comparison.

3. **Extract and Compare**

   * Click **"Extract BOM from DXF"**.
   * Click **"Compare BOM vs DXF"** to check for differences.

4. **Edit and View Tables**

   * Switch between tabs to view BOM from DXF, Excel, or both in split view.
   * Click on a cell to edit inline.
   * Changes are immediately flagged for synchronization.

5. **Drag and Drop Tags**

   * **Hold `CTRL`**, then:

     * **Click a row** in one table and **drag it** over a row in the other table.
     * Only the **P\&ID TAG** is transferred.
     * Releasing the mouse updates the target row.

6. **Copy and Paste (Full Row)**

   * Select a row in either table.
   * Press **`CTRL+C`** to copy.
   * Select a target row and press **`CTRL+V`** to overwrite it.
   * Useful for syncing between DXF and Excel or to an external spreadsheet.

7. **Zoom to Blocks in AutoCAD**

   * Select a row with a valid `P&ID TAG` and press **`CTRL+F`**.
   * AutoCAD will zoom to the block if open and found.
   * Alternatively, use **View > Zoom to Block in AutoCAD** from the menu.

8. **Save**

   * Use the File > Save menu to:

     * Save settings.
     * Save updated Excel BOM.
     * Save modified DXF file.

---

## ⌥️ Keyboard Shortcuts

| Shortcut      | Action                                             |
| ------------- | -------------------------------------------------- |
| `CTRL+Q`      | Compare BOM vs DXF immediately                     |
| `CTRL+C`      | Copy selected row                                  |
| `CTRL+V`      | Paste copied row into selected row                 |
| `CTRL+F`      | Zoom to selected block in AutoCAD                  |
| `CTRL+Z`      | Undo last modification to BOM                      |
| `CTRL` + Drag | Transfer `P&ID TAG` between tables (drag and drop) |

---

## 📄 Preparing Your DXF File

To work with this application, your DXF should meet these basic conditions:

* Components must be inserted as blocks with attributes.
* Each block must contain must have nearby a stricker containing P&ID Tag (e.g., `XV101`).
* Make sure AutoCAD is set up with meaningful tag data in block definitions.
* Avoid nested blocks or use simple insertions for consistent parsing.

---

## 🔒 License and Distribution

You may compile this Python project using `pyinstaller` to distribute as a standalone `.exe` file. If you plan to commercialize:

* Implement license key checks or cloud verification.
* Obfuscate with Cython or use code signing tools.

---

## 📧 Feedback & Contributions

Pull requests, suggestions, and issues are welcome! This project was developed with a focus on usability and integration with existing workflows.
