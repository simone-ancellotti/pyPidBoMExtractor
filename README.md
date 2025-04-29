# pyPidBoMExtractor

**pyPidBoMExtractor** is a powerful Python tool to extract, modify, and synchronize Bill of Materials (BOM) between DXF P&ID drawings and Excel spreadsheets.

It provides a simple user interface to:
- Extract BOMs from DXF files.
- Compare DXF BOMs with revised BOMs from Excel.
- Highlight missing or duplicated components.
- Import missing BOM entries into DXF.
- Edit BOMs manually with filtering and direct cell editing.
- Save updated BOMs into new Excel or DXF files.

---

## ✨ Features

- 📂 **Extract** BOM from DXF (AutoCAD) drawings.
- 📄 **Import** and **compare** with Excel BOMs.
- 🔍 **Highlight**:
  - Components missing in DXF or Excel.
  - Duplicated entries.
- 📥 **Import missing entries** automatically into Excel.
- ✏️ **Manual editing**:
  - Modify cell values.
  - Copy/paste between rows.
  - Drag & drop `P&ID TAG` from DXF to Excel BOM.
- 💾 **Save** updated DXF or Excel files.
- 🔄 **Undo** (planned future improvement).

---

## 🚀 How to Run

```bash
pip install -r requirements.txt
python bom_extractor_ui.py

You can also package it as an EXE using:
pyinstaller --onefile --icon=bom_valve_icon.ico --hidden-import=ezdxf --hidden-import=openpyxl --hidden-import=PIL bom_extractor_ui.py

## 🧰 How to Use

1. **Launch the App**  
   Run this command to start the graphical user interface:

   ```bash
   python bom_extractor_ui.py
   ```

2. **Upload Files**
   - **Upload DXF**: Click "Upload DXF" and choose your P&ID drawing.
   - **Upload Template Excel**: Click "Upload Template Excel" and select your BOM template.
   - **Upload Revised BOM Excel**: Optional, for comparison.

3. **Extract and Compare**
   - Click **"Extract BOM from DXF"**.
   - Click **"Compare BOM vs DXF"** to check for differences.

4. **Edit and View Tables**
   - Switch between tabs to view BOM from DXF, Excel, or both in split view.
   - Click on a cell to edit inline.
   - Changes are immediately flagged for synchronization.

5. **Drag and Drop Tags**
   - **Hold `CTRL`**, then:
     - **Click a row** in one table and **drag it** over a row in the other table.
     - Only the **P&ID TAG** is transferred.
     - Releasing the mouse updates the target row.

6. **Copy and Paste (Full Row)**
   - Select a row in either table.
   - Press **`CTRL+C`** to copy.
   - Select a target row and press **`CTRL+V`** to overwrite it.
   - Useful for syncing between DXF and Excel or to an external spreadsheet.

7. **Save**
   - Use the File > Save menu to:
     - Save settings.
     - Save updated Excel BOM.
     - Save modified DXF file.

---

## ⌥️ Keyboard Shortcuts

| Shortcut    | Action                                    |
|-------------|-------------------------------------------|
| `CTRL+Q`    | Compare BOM vs DXF immediately            |
| `CTRL+C`    | Copy selected row                         |
| `CTRL+V`    | Paste copied row into selected row        |
| `CTRL` + Drag | Transfer `P&ID TAG` between tables     |

