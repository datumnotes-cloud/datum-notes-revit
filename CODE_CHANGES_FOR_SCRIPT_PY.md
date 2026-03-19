# EXACT CODE CHANGES FOR script.py

# Copy-paste these sections into your script.py file

# ============================================================================

# SECTION 1: ADD IMPORTS (after line ~20, with other imports)

# ============================================================================

# Add this block after the existing imports:

try:
from excel_import import (
create_excel_template,
save_excel_template,
parse_excel_file,
validate_excel_file,
excel_available
)
EXCEL_SUPPORT = True
except ImportError:
EXCEL_SUPPORT = False # Dummy function when openpyxl not available
def excel_available():
return False

# ============================================================================

# SECTION 2: UPDATE build_ai_template() - MAKE INSTRUCTIONS STRICTER

# ============================================================================

# REPLACE the existing build_ai_template() with this:

def build_ai_template(project_name, rooms):
date_text = datetime.datetime.now().strftime("%Y-%m-%d")
return """IMPORTANT: Follow this format EXACTLY. Do not vary from the structure below.

Parse the meeting transcript and fill in this template. Keep each line following the exact pipe-delimited format.
Do NOT add explanations, extra text, or headers outside the data sections.

Room names: Use exact room identifiers. If unsure which room, use: UNASSIGNED

PROJECT: %s
DATE: %s

ACTION ITEMS: room | note | assigned to | due date
DECISIONS: room | note
QUESTIONS: room | note | assigned to
OBSERVATIONS: room | note
UNASSIGNED: note | category

EXAMPLE ENTRIES (FOLLOW THIS FORMAT EXACTLY):
ACTION ITEMS: Conference Room A | Finalize HVAC routing | John Smith | 2025-03-25
DECISIONS: Level 3 | Confirm concrete strength specification |
QUESTIONS: UNASSIGNED | Client approval on material selections | Project Manager
OBSERVATIONS: Room 201 | Electrical conduit route needs adjustment |
""" % (project_name, date_text)

# ============================================================================

# SECTION 3: ADD TWO NEW BUTTON HANDLERS TO RedlineWindow CLASS

# ============================================================================

# Add these methods AFTER the existing on_import_from_ai() method:

def on_download_excel_template(self, sender, args):
"""Handler for 'Download Template (Excel)' button"""
try: # Get project metadata
meta = get_project_metadata(self.doc)

        # Get existing room names
        rooms = sorted(list(set([
            _safe_text(_normalize_note(n).get("roomName", "")) or
            _safe_text(_normalize_note(n).get("roomDisplay", ""))
            for n in self._active_notes()
            if _safe_text(_normalize_note(n).get("roomName", "")) or
               _safe_text(_normalize_note(n).get("roomDisplay", ""))
        ])))

        # Create Excel workbook
        wb = create_excel_template(meta.get("projectName", "Untitled Project"), rooms)
        if not wb:
            forms.alert("Excel support not available. Please install openpyxl.\n\npip install openpyxl\n\nOtherwise, use the text-based import.", warn_icon=True)
            return

        # Show save file dialog
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
                    "Template saved successfully!\n\n%s\n\n"
                    "Next steps:\n"
                    "1. Open the Excel file\n"
                    "2. Follow the instructions in the worksheet\n"
                    "3. Fill in your data (or share with AI to complete)\n"
                    "4. Save the file\n"
                    "5. Come back and click 'Import from Excel'" % save_dialog.FileName
                )
            else:
                forms.alert("Failed to save Excel template.", warn_icon=True)
    except Exception as ex:
        forms.alert("Error creating Excel template:\n%s" % str(ex), warn_icon=True)

def on_import_from_excel(self, sender, args):
"""Handler for 'Import from Excel' button"""
try:
if not excel_available():
forms.alert(
"Excel support not available.\n\n"
"To enable Excel import:\n"
"1. Open PowerShell/Command Prompt\n"
"2. Run: pip install openpyxl\n"
"3. Restart Revit\n\n"
"For now, use the text-based 'Import from AI (Text)' option.",
warn_icon=True
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

        # Validate Excel file structure
        is_valid, error_msg = validate_excel_file(filepath)
        if not is_valid:
            forms.alert("Excel validation failed:\n%s" % error_msg, warn_icon=True)
            return

        # Parse Excel file
        parsed, parse_error = parse_excel_file(
            filepath,
            _normalize_category,
            self._match_room_from_text,
            unassigned_room_bucket
        )

        if parse_error:
            forms.alert("Error parsing Excel file:\n%s" % parse_error, warn_icon=True)
            return

        if not parsed or len(parsed) == 0:
            forms.alert("No valid entries found in Excel file.", warn_icon=True)
            return

        # Create notes from parsed items
        added = 0
        now = datetime.datetime.now()

        for item in parsed:
            # Determine room (handle UNASSIGNED)
            if bool(item.get("isUnassigned", False)):
                room = unassigned_room_bucket()
            else:
                room = self._match_room_from_text(item.get("room", ""))

            if not room:
                continue  # Room not found, skip this item

            # Build note object
            note = {
                "id": self._make_note_id(),
                "timestamp": now.strftime("%Y-%m-%d %H:%M:%S"),
                "roomId": room.get("roomId", ""),
                "roomDisplay": room.get("roomDisplay", ""),
                "roomNumber": room.get("number", ""),
                "roomName": room.get("name", ""),
                "level": room.get("level", ""),
                "elementId": room.get("elementId", ""),
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

        forms.alert("Successfully imported %d notes from Excel\n\nFile: %s" % (added, os.path.basename(filepath)))

    except Exception as ex:
        forms.alert("Error importing from Excel:\n%s" % str(ex), warn_icon=True)

# ============================================================================

# SECTION 4: RENAME EXISTING METHOD (for clarity)

# ============================================================================

# RENAME the existing on_import_from_ai() method to:

# on_import_from_ai_text()

# (This keeps the old text-based import but makes it clear it's text-based)

# If you want to keep the same name, that's fine too - just wire both text and

# Excel buttons to their respective handlers.

# ============================================================================

# SECTION 5: WIRE UP BUTTONS IN \_wire_events() METHOD

# ============================================================================

# Find the \_wire_events() method and add these lines where buttons are wired:

# After the existing:

# self.copyAiTemplateButton.Click += self.on_copy_ai_template

# self.importAiButton.Click += self.on_import_from_ai

# Add:

    # Wire Excel buttons if support is available
    if EXCEL_SUPPORT:
        self.downloadExcelButton.Click += self.on_download_excel_template
        self.importExcelButton.Click += self.on_import_from_excel
    else:
        # Disable Excel buttons if openpyxl not available
        if hasattr(self, 'downloadExcelButton'):
            self.downloadExcelButton.IsEnabled = False
            self.downloadExcelButton.ToolTip = "openpyxl not installed"
        if hasattr(self, 'importExcelButton'):
            self.importExcelButton.IsEnabled = False
            self.importExcelButton.ToolTip = "openpyxl not installed"

# ============================================================================

# SECTION 6: UPDATE redline_ui.xaml

# ============================================================================

# Find the footer WrapPanel with the AI template buttons, and REPLACE:

#

# <Button x:Name="copyAiTemplateButton" ... >Copy AI Template</Button>

# <Button x:Name="importAiButton" ... >Import from AI</Button>

#

# WITH:

#

# <!-- AI Template Options (Text) -->

# <Button x:Name="copyAiTemplateButton" ... Margin="0,0,5,0" ToolTip="Copy text template to clipboard">Copy Template (Text)</Button>

# <Button x:Name="importAiButton" ... ToolTip="Paste AI output in text format">Import from AI (Text)</Button>

#

# <!-- AI Template Options (Excel) -->

# <Button x:Name="downloadExcelButton" ... Margin="10,0,5,0" ToolTip="Download Excel template with instructions">Download Template (Excel)</Button>

# <Button x:Name="importExcelButton" ... ToolTip="Import completed Excel file from AI">Import from Excel</Button>

# Both sets of buttons can coexist for maximum flexibility.
