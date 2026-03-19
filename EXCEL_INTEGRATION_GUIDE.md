# Implementation Guide for integrating Excel Import into script.py

# Add these sections to your existing script.py file

# ============================================================================

# 1. ADD THIS IMPORT NEAR THE TOP (after other imports)

# ============================================================================

# At line ~20, after the existing imports, add:

"""
try:
from excel_import import (
create_excel_template, save_excel_template, parse_excel_file,
validate_excel_file, excel_available
)
EXCEL_SUPPORT = True
except ImportError:
EXCEL_SUPPORT = False
excel_available = lambda: False
"""

# ============================================================================

# 2. ADD THESE METHODS TO THE RedlineWindow CLASS

# ============================================================================

"""
def on_download_excel_template(self, sender, args):
'''Handler for "Download Excel Template" button'''
try: # Get project metadata and existing rooms
meta = get_project_metadata(self.doc)
rooms = sorted(list(set([
_safe_text(_normalize_note(n).get("roomName", "")) or
_safe_text(_normalize_note(n).get("roomDisplay", ""))
for n in self._active_notes()
])))

        # Create Excel workbook
        wb = create_excel_template(meta.get("projectName", "Untitled Project"), rooms)
        if not wb:
            forms.alert("Excel support not available. Please install openpyxl or use text-based import.")
            return

        # Show save dialog
        save_dialog = SaveFileDialog()
        save_dialog.FileName = "%s_Redline_Template_%s.xlsx" % (
            _sanitize_filename(meta.get("projectName", "project")),
            datetime.datetime.now().strftime("%Y%m%d")
        )
        save_dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        save_dialog.DefaultExt = "xlsx"

        if save_dialog.ShowDialog():
            if save_excel_template(wb, save_dialog.FileName):
                forms.alert(
                    "Template saved to:\\n%s\\n\\n"
                    "1. Open the file in Excel\\n"
                    "2. Fill in the data following the instructions\\n"
                    "3. Save the file\\n"
                    "4. Click 'Import from Excel' to import" % save_dialog.FileName
                )
            else:
                forms.alert("Failed to save Excel template.", warn_icon=True)
    except Exception as ex:
        forms.alert("Error creating Excel template: %s" % str(ex), warn_icon=True)

def on_import_from_excel(self, sender, args):
'''Handler for "Import from Excel" button'''
try:
if not excel_available():
forms.alert(
"Excel support not available.\\n\\n"
"To enable Excel import:\\n"
"1. Install openpyxl: pip install openpyxl\\n"
"2. Restart Revit/pyRevit\\n\\n"
"For now, use 'Import from AI (Text)' with the text template."
)
return

        # Show file open dialog
        open_dialog = OpenFileDialog()
        open_dialog.FileName = ""
        open_dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        open_dialog.DefaultExt = "xlsx"

        if not open_dialog.ShowDialog():
            return  # User cancelled

        filepath = open_dialog.FileName

        # Validate Excel file
        is_valid, error_msg = validate_excel_file(filepath)
        if not is_valid:
            forms.alert("Excel validation failed:\\n%s" % error_msg, warn_icon=True)
            return

        # Parse Excel file
        parsed, parse_error = parse_excel_file(
            filepath,
            _normalize_category,
            self._match_room_from_text,
            unassigned_room_bucket
        )

        if parse_error:
            forms.alert("Error parsing Excel file:\\n%s" % parse_error, warn_icon=True)
            return

        if not parsed:
            forms.alert("No valid entries found in Excel file.", warn_icon=True)
            return

        # Import parsed items
        added = 0
        now = datetime.datetime.now()

        for item in parsed:
            # Determine room
            if bool(item.get("isUnassigned", False)):
                room = unassigned_room_bucket()
            else:
                room = self._match_room_from_text(item.get("room", ""))

            if not room:
                continue  # Room not found, skip

            # Create note object
            note = {
                "id": self._make_note_id(),
                "timestamp": now.strftime("%Y-%m-%d %H:%M:%S"),
                "roomId": room["roomId"],
                "roomDisplay": room["roomDisplay"],
                "roomNumber": room["number"],
                "roomName": room["name"],
                "level": room["level"],
                "elementId": room["elementId"],
                "text": _safe_text(item.get("text", "")).strip(),
                "completed": False,
                "completedAt": "",
                "completedBy": "",
                "editedAt": "",
                "editedBy": "",
                "category": _normalize_category(item.get("category", "Observation")),
                "assignedTo": _safe_text(item.get("assignedTo", "")).strip(),
                "dueDate": _safe_text(item.get("dueDate", "")).strip(),
                "deleted": False,
                "deletedAt": "",
                "imported": True,
                "duplicateFrom": "",
                "comments": []
            }

            if note["text"]:
                self.all_notes.append(note)
                added += 1

        # Save and refresh UI
        self._save()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()

        forms.alert("✓ Imported %d notes from Excel\\n\\nFile: %s" % (added, os.path.basename(filepath)))

    except Exception as ex:
        forms.alert("Error importing from Excel: %s" % str(ex), warn_icon=True)

def on_import_from_ai_text(self, sender, args):
'''Updated handler for text-based import (renamed from on_import_from_ai)''' # Keep all the existing text-based import logic here # This becomes a fallback option
pass
"""

# ============================================================================

# 3. WIRE UP BUTTONS IN **init** (in \_wire_events method)

# ============================================================================

"""
Add these lines in \_wire_events() after existing button wirings:

    if EXCEL_SUPPORT:
        self.downloadExcelButton.Click += self.on_download_excel_template
        self.importExcelButton.Click += self.on_import_from_excel

    # Rename existing button handler
    self.importAiButton.Click += self.on_import_from_ai_text

"""

# ============================================================================

# 4. UPDATE REDLINE_UI.XAML BUTTONS

# ============================================================================

"""
In the footer WrapPanel where the AI template buttons are, replace:

    <Button x:Name="copyAiTemplateButton" ...>Copy AI Template</Button>
    <Button x:Name="importAiButton" ...>Import from AI</Button>

With:

    <!-- AI Template Options -->
    <Button x:Name="copyAiTemplateButton" ...>Copy Template (Text)</Button>
    <Button x:Name="downloadExcelButton" ...>Download Template (Excel)</Button>
    <Button x:Name="importAiButton" ...>Import from AI (Text)</Button>
    <Button x:Name="importExcelButton" ...>Import from Excel</Button>

"""

# ============================================================================

# 5. RECOMMENDED: ADD STRICTER TEXT-BASED INSTRUCTIONS

# ============================================================================

"""
Update build_ai_template() to have clearer, stricter instructions:

def build_ai_template(project_name, rooms):
date_text = datetime.datetime.now().strftime("%Y-%m-%d")
return \"\"\"IMPORTANT: Follow this format EXACTLY and do not vary from it.

Parse the meeting transcript below and fill in this template. Each line must follow the exact format.
Do NOT add extra text, headers, or explanations outside of the data.
For room identification: use exact room names/numbers. If unsure, use: UNASSIGNED

PROJECT: %s
DATE: %s

ACTION ITEMS: room | note text | assigned to | due date
DECISIONS: room | note text
QUESTIONS: room | note text | assigned to
OBSERVATIONS: room | note text
UNASSIGNED: note text | category

Example:
ACTION ITEMS: Conference Room A | Review structural design | John Smith | 2025-03-25
DECISIONS: Level 2 | MEP routing finalized |
QUESTIONS: UNASSIGNED | Confirm client preferences | Project Manager
\"\"\" % (project_name, date_text)
"""

# ============================================================================

# 6. VALIDATION HELPER (for text-based import)

# ============================================================================

"""
Add this validation function to improve text-based import strictness:

def validate_text_import(text):
'''
Validate that pasted text follows proper template format.
Returns: (is_valid, error_message)
'''
required_sections = ["ACTION ITEMS:", "DECISIONS:", "QUESTIONS:", "OBSERVATIONS:"]
lines = text.strip().split("\\n")

    found_sections = set()
    for line in lines:
        for section in required_sections:
            if section in line:
                found_sections.add(section)

    if len(found_sections) == 0:
        return False, "No valid sections found (ACTION ITEMS, DECISIONS, etc.)"

    # Check for pipe delimiter usage
    has_pipes = any("|" in line for line in lines if line.strip() and not ":" in line)
    if not has_pipes:
        return False, "Format error: Lines should use | to separate fields"

    return True, ""

"""
