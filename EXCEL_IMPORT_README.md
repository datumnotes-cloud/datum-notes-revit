# **Excel Import Feature for Datum Notes - Implementation Summary**

## **What Changed**

Your AI import workflow now has **two paths**:

1. **Text-Based (Existing)** — Works like before, but with stricter instructions
2. **Excel-Based (New)** — Download template → AI fills it → Upload Excel file

---

## **New Features**

### **1. Download Excel Template**

- User clicks "Download Template (Excel)"
- Browser/dialog saves a formatted `.xlsx` file
- File includes:
  - **Instructions sheet** with strict formatting rules for AI
  - **Data columns**: Room | Note | Category | Assigned To | Due Date
  - **Example row** showing proper format
  - **Reference sheet** with available rooms for auto-complete

### **2. Import from Excel**

- User clicks "Import from Excel"
- Select the completed Excel file
- System automatically:
  - Validates all required fields (Room, Note, Category)
  - Validates category values (must be: Action Item, Decision, Question, Observation)
  - Validates date format (YYYY-MM-DD)
  - Shows clear error messages for any issues
  - Creates notes in Revit with all metadata

### **3. Improved Text-Based Import**

- "Copy Template (Text)" button — same as before
- "Import from AI (Text)" button — same workflow, but stricter instructions in template

---

## **Why This Is Better**

| Aspect              | Text-Based                        | Excel-Based                             |
| ------------------- | --------------------------------- | --------------------------------------- |
| **Format Control**  | Pipe-delimited, easy to mess up   | Structured columns, hard to break       |
| **Verification**    | Hard to spot errors               | Visual verification before upload       |
| **AI Instructions** | Implicit (AI has to guess format) | Explicit instructions sheet in workbook |
| **Error Messages**  | Generic                           | Specific row numbers & field issues     |
| **Validation**      | Manual                            | Automatic with feedback                 |
| **Team Handoff**    | Copy-paste text                   | Send Excel file                         |

---

## **Files Created**

1. **`excel_import.py`** — New module with:
   - `create_excel_template()` — Generates workbook
   - `save_excel_template()` — Saves to disk
   - `parse_excel_file()` — Reads completed Excel
   - `validate_excel_file()` — Validates structure
   - `excel_available()` — Checks if openpyxl is installed

2. **`EXCEL_INTEGRATION_GUIDE.md`** — Detailed code snippets for integration

---

## **Installation Steps**

### **Step 1: Install openpyxl**

Open PowerShell/Command Prompt and run:

```powershell
pip install openpyxl
```

Or for IronPython (if pyRevit doesn't use system Python):

```
C:\path\to\python.exe -m pip install openpyxl
```

### **Step 2: Copy excel_import.py**

Place the new `excel_import.py` file in the same folder as `script.py`:

```
DatumNotes.extension\Datum Notes.tab\Project Tools.panel\Meeting Notes.pushbutton\
```

### **Step 3: Update script.py**

Add to imports (line ~20):

```python
try:
    from excel_import import (
        create_excel_template, save_excel_template, parse_excel_file,
        validate_excel_file, excel_available
    )
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False
    excel_available = lambda: False
```

### **Step 4: Add Event Handlers**

Copy the two new methods from the integration guide:

- `on_download_excel_template()`
- `on_import_from_excel()`

And add to `_wire_events()`:

```python
if EXCEL_SUPPORT:
    self.downloadExcelButton.Click += self.on_download_excel_template
    self.importExcelButton.Click += self.on_import_from_excel
```

### **Step 5: Update XAML**

Add two buttons to the footer in `redline_ui.xaml`:

```xml
<Button x:Name="downloadExcelButton" ...>Download Template (Excel)</Button>
<Button x:Name="importExcelButton" ...>Import from Excel</Button>
```

---

## **How AI Instructions Work (Excel)**

The generated Excel template includes a dedicated **"INSTRUCTIONS FOR AI ASSISTANT"** sheet with:

✓ **Explicit rules** on Room identification (exact names, UNASSIGNED format)  
✓ **Field definitions** for each column (what goes in Note vs Category)  
✓ **Valid category values** (Action Item, Decision, Question, Observation)  
✓ **Optional field handling** (Assigned To, Due Date)  
✓ **DO NOT list** (don't delete headers, don't merge cells, etc.)  
✓ **Example row** showing completed data

This eliminates the guesswork for AI assistants. They see **exactly** what format to use.

---

## **Fallback: If openpyxl Not Available**

If `pip install openpyxl` fails:

- Excel buttons will be disabled
- Text-based import remains available
- Users can still use "Copy Template (Text)" and "Import from AI (Text)"

The code gracefully handles missing openpyxl:

```python
try:
    from openpyxl import ...
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False
```

---

## **Example Workflow**

**Before (Text-Based):**

1. Click "Copy AI Template" → Text copied to clipboard
2. Paste into ChatGPT with transcript
3. Copy AI output
4. Click "Import from AI"
5. Paste text
6. Hope it parses correctly ❌

**After (Excel-Based):**

1. Click "Download Template (Excel)" → Save .xlsx file
2. Open in Excel, paste data, or send to AI to fill
3. Save the file
4. Click "Import from Excel"
5. Select file
6. System validates and imports automatically ✅

---

## **Summary of AI Instruction Improvements**

### **Old Text Template:**

```
Parse the transcript below and fill in this template exactly...
PROJECT: My Project
DATE: 2025-03-19
ACTION ITEMS: room | note | assigned to | due date
...
```

### **New Excel Template (Strengths):**

- **Explicit instructions** — Full page of formatting rules
- **Visual structure** — Columns with headers, not text parsing
- **Example row** — Colored, shows exactly what good data looks like
- **Field validation** — System checks category values, dates, required fields
- **Error feedback** — "Row 5: Category must be 'Action Item', not 'Actionitem'"
- **Room reference sheet** — Available rooms listed (for context)

---

## **Testing Checklist**

- [ ] openpyxl installed (`pip install openpyxl`)
- [ ] `excel_import.py` in Meeting Notes button folder
- [ ] `script.py` updated with imports and handlers
- [ ] `redline_ui.xaml` has new buttons
- [ ] Download template → Opens Excel dialog
- [ ] Excel file downloads with proper formatting
- [ ] Excel instructions are clear and complete
- [ ] Can fill Excel manually and import successfully
- [ ] Validation catches missing required fields
- [ ] Validation catches invalid categories
- [ ] Validation catches invalid dates (malformed)
- [ ] Error messages are helpful
- [ ] Text-based import still works as fallback

---

## **Next Steps**

1. ✅ Review this summary and the generated `excel_import.py` module
2. ⏳ Install openpyxl: `pip install openpyxl`
3. ⏳ Integrate code into `script.py` using the guide
4. ⏳ Update `redline_ui.xaml` with new buttons
5. ⏳ Test with sample Excel files
6. ⏳ Refine AI prompt in template based on testing
7. ⏳ Deploy to users

---

## **Questions & Customization**

**Q: Can I modify the Excel template structure?**  
A: Yes! The `create_excel_template()` function is fully customizable. Change column widths, colors, instructions, etc.

**Q: What if users don't have Excel?**  
A: They can open .xlsx files in Google Sheets, LibreOffice, or any spreadsheet app. Or use the text-based import.

**Q: Can I change the AI instructions in the template?**  
A: Yes! Edit the `instructions` list in `create_excel_template()` function (around line 130).

**Q: What happens if a room doesn't match?**  
A: That row is skipped. Add fuzzy matching in `_match_room_from_text()` if needed.

---

**Questions? Issues? Let me know and I'll help refine the implementation!**
