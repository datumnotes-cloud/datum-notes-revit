# -*- coding: utf-8 -*-
from __future__ import print_function

import os
import re
import json
import codecs
import base64
import datetime
import traceback

from pyrevit import revit, DB, forms, script

from System.Windows import Thickness, Visibility, VerticalAlignment, TextWrapping, FontWeights, HorizontalAlignment, GridLength
from System.Windows.Controls import TextBlock, StackPanel, CheckBox, Border, Orientation, Button, Grid, ColumnDefinition, Expander, TextBox
from System.Diagnostics import Process
from Microsoft.Win32 import SaveFileDialog, OpenFileDialog
from System.Windows.Media import BrushConverter
from System.Collections.Generic import List
from System.Windows import Clipboard
from System.Windows.Input import Key, ModifierKeys, Keyboard


__persistentengine__ = True


UPLOAD_INFO_URL = "https://datumnotes.com/from-revit"
UPDATE_INFO_URL = "https://datumnotes.com/revit/releases"
TUTORIAL_URL = "https://datumnotes.com/revit/tutorial"
EXTENSION_VERSION = "0.3.0"
CATEGORY_OPTIONS = ["Decision", "Action Item", "Question", "Observation"]
BADGE_COLORS = {
    "Decision": "#2563EB",
    "Action Item": "#DC2626",
    "Question": "#F59E0B",
    "Observation": "#6B7280"
}
DEFAULT_TEAM_MEMBERS = ["PM", "Architect", "Coordinator", "MEP", "Owner"]
SORT_OPTIONS = ["Newest First", "Oldest First", "By Room"]
CATEGORY_TAB_OPTIONS = ["All", "Action Item", "Question", "Decision", "Observation", "Pending"]
UNASSIGNED_ROOM_ID = "UNASSIGNED"
UNASSIGNED_ROOM_DISPLAY = "GENERAL | General"
ICON_PNG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAYAAADimHc4AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAHjSURBVHhe7dExbsNQEANR9+5ygpw8t0ydNHJDwAB3IYZQNMVr7IX9iXk8Pz5/0PPQD/C3CFBGgDIClBGgjABlBCgjQBkByghQRoAyApQRoOy0AN9fz1vR/VsEWNL9WwRY0v1bBFjS/VuxAPr91aX2EcCU2kcAU2ofAUypfQQwpfYRwJTaRwBTah8BTKl9BDCl9hHAlNpXC6D3bfo+Nb13EeCg71PTexcBDvo+Nb13EeCg71PTe1ctwNWk9hHAlNpHAFNqHwFMqX0EMKX2EcCU2kcAU2ofAUypfbUAet+m71PTexcBDvo+Nb13EeCg71PTexcBDvo+Nb131QJcTWofAUypfQQwpfYRwJTaRwBTah8BTKl9BDCl9hHAlNpXC6D3afr/U2f/3gsBTGf/3svtAujn+v0703sXAeT7d6b3rlqAFn2n+97pvYsA5nun9y4CmO+d3rsIYL53eu+6XYCt1D4CmFL7CGBK7SOAKbWPAKbUPgKYUvsIYErtI4AptY8AptQ+AphS+whgSu0jgCm1jwCm1D4CmFL7CGBK7SOAKbWPAKbUPgKYUvtiAf473b9FgCXdv0WAJd2/RYAl3b91WgDsEKCMAGUEKCNAGQHKCFBGgDIClBGgjABlBCgjQBkByn4BQpL38pWaWnMAAAAASUVORK5CYII="

THEME_WINDOW_BG = "#262626"
THEME_PANEL_BG = "#303030"
THEME_PANEL_BORDER = "#4A4A4A"
THEME_CARD_BG = "#383838"
THEME_INPUT_BG = "#3A3A3A"
THEME_BUTTON_BG = "#444444"
THEME_BUTTON_BORDER = "#606060"


def _brush(color_hex):
    return BrushConverter().ConvertFromString(color_hex)


def _safe_text(value):
    if value is None:
        return ""
    try:
        return str(value)
    except Exception:
        return ""


def _sanitize_filename(name):
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", _safe_text(name).strip())
    return cleaned.strip("_") or "revit_project"


def _normalize_category(value):
    text = _safe_text(value).strip()
    return text if text in CATEGORY_OPTIONS else "Observation"


def _lookup_parameter_text(element, names):
    if not element:
        return ""

    for name in names:
        try:
            param = element.LookupParameter(name)
            if not param:
                continue

            val = ""
            if hasattr(param, "AsString"):
                val = _safe_text(param.AsString())
            if not val and hasattr(param, "AsValueString"):
                val = _safe_text(param.AsValueString())

            if val and val.strip():
                return val.strip()
        except Exception:
            pass

    return ""


def get_project_metadata(doc):
    project_info = getattr(doc, "ProjectInformation", None)

    project_name = _safe_text(getattr(project_info, "Name", "")).strip() or _safe_text(getattr(doc, "Title", "")).strip() or "Untitled Project"
    architect_name = _lookup_parameter_text(project_info, ["Architect", "Architect Name", "Organization Name", "Company", "Author"]) or "Not specified"
    project_address = _lookup_parameter_text(project_info, ["Project Address", "Address", "Building Address", "Site Address"])

    return {
        "projectName": project_name,
        "architectName": architect_name,
        "projectAddress": project_address
    }


def get_current_user_name(doc):
    try:
        app = getattr(doc, "Application", None)
        username = _safe_text(getattr(app, "Username", "")).strip()
        if username:
            return username
    except Exception:
        pass

    try:
        env_user = _safe_text(os.environ.get("USERNAME", "")).strip()
        if env_user:
            return env_user
    except Exception:
        pass

    return "Unknown User"


def ensure_icon_png():
    try:
        bundle_script = script.get_bundle_file("script.py")
        if not bundle_script:
            return

        icon_path = os.path.join(os.path.dirname(bundle_script), "icon.png")
        if os.path.exists(icon_path):
            return

        raw = base64.b64decode(ICON_PNG_BASE64)
        with open(icon_path, "wb") as icon_file:
            icon_file.write(raw)
    except Exception:
        pass


def _normalize_note(note):
    base = note if isinstance(note, dict) else {}
    category = _normalize_category(base.get("category", "Observation"))
    comments = base.get("comments", [])

    # Prefer assigned_to when present to handle legacy/imported payloads
    # where assignedTo and assigned_to can conflict.
    assigned_value = _safe_text(
        base.get("assigned_to")
        or base.get("assignedTo")
        or base.get("assigned")
        or base.get("assignee")
    ).strip()

    room_id = _safe_text(base.get("roomId", "")).strip()
    room_display = _safe_text(base.get("roomDisplay", "")).strip()
    room_number = _safe_text(base.get("roomNumber", "")).strip()
    room_name = _safe_text(base.get("roomName", "")).strip()
    room_level = _safe_text(base.get("level", "")).strip()

    if room_id == UNASSIGNED_ROOM_ID or room_display.upper().startswith("UNASSIGNED"):
        room_display = UNASSIGNED_ROOM_DISPLAY
        room_number = "GENERAL"
        room_name = "General"
        room_level = "General"

    if not room_display:
        room_display = "Unknown Room"

    normalized_comments = []
    if isinstance(comments, list):
        for c in comments:
            if not isinstance(c, dict):
                continue
            txt = _safe_text(c.get("text", "")).strip()
            if not txt:
                continue
            normalized_comments.append({
                "timestamp": _safe_text(c.get("timestamp", "")),
                "author": _safe_text(c.get("author", "")),
                "text": txt
            })

    return {
        "id": _safe_text(base.get("id")),
        "timestamp": _safe_text(base.get("timestamp")),
        "roomId": room_id,
        "roomNumber": room_number,
        "roomName": room_name,
        "roomDisplay": room_display,
        "level": room_level,
        "elementId": _safe_text(base.get("elementId")),
        "text": _safe_text(base.get("text")),
        "completed": bool(base.get("completed", False)),
        "completedAt": _safe_text(base.get("completedAt")),
        "completedBy": _safe_text(base.get("completedBy")),
        "pending": bool(base.get("pending", False)),
        "editedAt": _safe_text(base.get("editedAt")),
        "editedBy": _safe_text(base.get("editedBy")),
        "category": category,
        "assignedTo": assigned_value,
        "dueDate": _safe_text(base.get("dueDate")) if category == "Action Item" else "",
        "deleted": bool(base.get("deleted", False)),
        "deletedAt": _safe_text(base.get("deletedAt")),
        "imported": bool(base.get("imported", False)),
        "duplicateFrom": _safe_text(base.get("duplicateFrom")),
        "comments": normalized_comments
    }


def _normalize_work_item(item):
    base = item if isinstance(item, dict) else {}
    timestamp = _safe_text(base.get("timestamp", "")).strip()
    parsed = _parse_note_datetime(timestamp)
    day_text = _safe_text(base.get("day", "")).strip()
    if not day_text and parsed != datetime.datetime.min:
        day_text = parsed.strftime("%Y-%m-%d")
    if not day_text and timestamp:
        day_text = timestamp[:10]

    return {
        "id": _safe_text(base.get("id")),
        "timestamp": timestamp,
        "day": day_text,
        "text": _safe_text(base.get("text", "")).strip(),
        "author": _safe_text(base.get("author", "")).strip(),
        "editedAt": _safe_text(base.get("editedAt", "")).strip(),
        "editedBy": _safe_text(base.get("editedBy", "")).strip()
    }


def _parse_note_datetime(value):
    text = _safe_text(value).strip()
    if not text:
        return datetime.datetime.min

    for pattern in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y"]:
        try:
            return datetime.datetime.strptime(text, pattern)
        except Exception:
            pass

    return datetime.datetime.min


def _ensure_datum_notes_folder():
    documents_dir = os.path.join(os.path.expanduser("~"), "Documents")
    fallback_dir = os.path.join(documents_dir, "DatumNotes")

    try:
        if not os.path.exists(fallback_dir):
            os.makedirs(fallback_dir)
    except Exception:
        return documents_dir

    return fallback_dir


def redline_config_path():
    return os.path.join(_ensure_datum_notes_folder(), "redline_config.json")


def load_redline_config():
    default_config = {
        "customAssignees": [],
        "manualResolvedRooms": [],
        "lastCustomAssignee": ""
    }

    path = redline_config_path()
    if not os.path.exists(path):
        return default_config

    try:
        with codecs.open(path, "r", "utf-8") as handle:
            data = json.load(handle)
        if not isinstance(data, dict):
            return default_config

        out = dict(default_config)
        custom_vals = data.get("customAssignees", [])
        resolved_vals = data.get("manualResolvedRooms", [])
        last_custom = _safe_text(data.get("lastCustomAssignee", "")).strip()
        out["customAssignees"] = [x for x in custom_vals if _safe_text(x).strip()] if isinstance(custom_vals, list) else []
        out["manualResolvedRooms"] = [x for x in resolved_vals if _safe_text(x).strip()] if isinstance(resolved_vals, list) else []
        out["lastCustomAssignee"] = last_custom
        return out
    except Exception:
        return default_config


def save_redline_config(config):
    try:
        with codecs.open(redline_config_path(), "w", "utf-8") as handle:
            json.dump(config, handle, indent=2)
        return True
    except Exception:
        return False


def _should_use_documents_fallback(doc, model_path):
    if not model_path:
        return True

    lowered = model_path.strip().lower()
    if lowered.startswith("bim 360://") or lowered.startswith("autodesk docs://") or lowered.startswith("cloud://"):
        return True

    try:
        if hasattr(doc, "IsModelInCloud") and doc.IsModelInCloud:
            return True
    except Exception:
        pass

    return False


def _room_number(room):
    number = _safe_text(getattr(room, "Number", "")).strip()
    if number:
        return number

    param = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER)
    if param:
        return _safe_text(param.AsString()).strip()

    return ""


def _room_name(room):
    param = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
    if param:
        return _safe_text(param.AsString()).strip()
    return ""


def collect_rooms(doc):
    groups = {}

    collector = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_Rooms) \
        .WhereElementIsNotElementType()

    for room in collector:
        try:
            if hasattr(room, "Area") and room.Area <= 0:
                continue
        except Exception:
            pass

        level_name = "Unassigned Level"
        try:
            level = doc.GetElement(room.LevelId)
            if level and getattr(level, "Name", None):
                level_name = level.Name
        except Exception:
            pass

        number = _room_number(room) or "?"
        name = _room_name(room) or "Unnamed Room"

        item = {
            "roomId": _safe_text(room.UniqueId),
            "elementId": _safe_text(room.Id.IntegerValue),
            "roomDisplay": "%s | %s - %s" % (level_name, number, name),
            "level": level_name,
            "number": number,
            "name": name
        }
        groups.setdefault(level_name, []).append(item)

    room_items = []
    for lvl in sorted(groups.keys()):
        rooms = sorted(groups[lvl], key=lambda x: (x["number"], x["name"]))
        room_items.extend(rooms)

    return room_items


def unassigned_room_bucket():
    return {
        "roomId": UNASSIGNED_ROOM_ID,
        "elementId": "",
        "roomDisplay": UNASSIGNED_ROOM_DISPLAY,
        "level": "General",
        "number": "GENERAL",
        "name": "General"
    }


def json_path_for_document(doc):
    model_path = _safe_text(getattr(doc, "PathName", ""))
    if _should_use_documents_fallback(doc, model_path):
        folder = _ensure_datum_notes_folder()
        project_name = _safe_text(getattr(doc, "Title", "")) or "revit_project"
    else:
        folder = os.path.dirname(model_path)
        project_name = os.path.splitext(os.path.basename(model_path))[0]

    filename = "%s_redline.json" % _sanitize_filename(project_name)
    return os.path.join(folder, filename)


def html_path_for_document(doc):
    model_path = _safe_text(getattr(doc, "PathName", ""))
    if _should_use_documents_fallback(doc, model_path):
        folder = _ensure_datum_notes_folder()
        project_name = _safe_text(getattr(doc, "Title", "")) or "revit_project"
    else:
        folder = os.path.dirname(model_path)
        project_name = os.path.splitext(os.path.basename(model_path))[0]

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "%s_redline_%s.html" % (_sanitize_filename(project_name), timestamp)
    return os.path.join(folder, filename)


def load_project_payload(path):
    if not path or not os.path.exists(path):
        return None

    try:
        with codecs.open(path, "r", "utf-8") as handle:
            return json.load(handle)
    except Exception:
        return None


def load_notes(path):
    data = load_project_payload(path)
    if data is None:
        return []

    if isinstance(data, dict):
        notes = data.get("redlineItems", data.get("notes", []))
        comments_by_note = data.get("commentsByNoteId", {})
        if isinstance(notes, list):
            out = []
            for note in notes:
                n = _normalize_note(note)
                note_id = _safe_text(n.get("id"))
                if isinstance(comments_by_note, dict) and note_id and isinstance(comments_by_note.get(note_id), list):
                    n["comments"] = _normalize_note({"comments": comments_by_note.get(note_id, [])}).get("comments", [])
                out.append(n)
            return out
        return []

    if isinstance(data, list):
        return [_normalize_note(note) for note in data]

    return []


def load_work_log_items(path):
    data = load_project_payload(path)
    if not isinstance(data, dict):
        return []

    items = data.get("workLogItems", data.get("workItems", []))
    if not isinstance(items, list):
        return []

    normalized = []
    for item in items:
        work_item = _normalize_work_item(item)
        if work_item.get("text"):
            normalized.append(work_item)
    return normalized


def save_notes(path, doc_title, notes, work_log_items=None, custom_assignees=None):
    if not path:
        return False

    comments_by_note = {}
    normalized_notes = []
    for note in notes:
        n = _normalize_note(note)
        normalized_notes.append(n)
        nid = _safe_text(n.get("id"))
        if nid:
            comments_by_note[nid] = n.get("comments", [])

    normalized_work_items = []
    for item in work_log_items or []:
        normalized_item = _normalize_work_item(item)
        if normalized_item.get("text"):
            normalized_work_items.append(normalized_item)

    payload = {
        "redlineProject": _safe_text(doc_title),
        "redlineSavedAt": datetime.datetime.now().isoformat(),
        "redlineItems": normalized_notes,
        "commentsByNoteId": comments_by_note,
        "workLogItems": normalized_work_items,
        "customAssignees": [_safe_text(x).strip() for x in (custom_assignees or []) if _safe_text(x).strip()]
    }

    try:
        with codecs.open(path, "w", "utf-8") as handle:
            json.dump(payload, handle, indent=2)
        return True
    except Exception:
        return False


def html_escape(value):
    text = _safe_text(value)
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = text.replace('"', "&quot;")
    return text


class ExportOptionsWindow(forms.WPFWindow):
    def __init__(self, xaml_path, notes):
        forms.WPFWindow.__init__(self, xaml_path)
        self.notes = [_normalize_note(n) for n in notes]
        self.result = None
        self._bind_filter_choices()
        self.okButton.Click += self.on_ok
        self.cancelButton.Click += self.on_cancel

    def _bind_filter_choices(self):
        categories = sorted(list(set([_normalize_category(n.get("category", "Observation")) for n in self.notes])))
        assigned = sorted(list(set([_safe_text(n.get("assignedTo", "")).strip() for n in self.notes if _safe_text(n.get("assignedTo", "")).strip()])))

        self.filterCategoryCombo.Items.Clear()
        self.filterCategoryCombo.Items.Add("All")
        for c in categories:
            self.filterCategoryCombo.Items.Add(c)
        self.filterCategoryCombo.SelectedItem = "All"

        self.filterAssignedCombo.Items.Clear()
        self.filterAssignedCombo.Items.Add("All")
        for a in assigned:
            self.filterAssignedCombo.Items.Add(a)
        self.filterAssignedCombo.SelectedItem = "All"

        self.sortOrderCombo.Items.Clear()
        self.sortOrderCombo.Items.Add("Newest First")
        self.sortOrderCombo.Items.Add("Oldest First")
        self.sortOrderCombo.SelectedItem = "Newest First"

    def _is_checked(self, box):
        return bool(getattr(box, "IsChecked", False))

    def on_ok(self, sender, args):
        include_completed = self._is_checked(self.includeCompletedCheck)
        include_uncompleted = self._is_checked(self.includeUncompletedCheck)
        if (not include_completed) and (not include_uncompleted):
            forms.alert("Select at least one of Completed items or Uncompleted items.", warn_icon=True)
            return

        self.result = {
            "include_project": self._is_checked(self.includeProjectCheck),
            "include_datetime": self._is_checked(self.includeDateCheck),
            "include_room_numbers": self._is_checked(self.includeRoomCheck),
            "include_assigned": self._is_checked(self.includeAssignedCheck),
            "include_due": self._is_checked(self.includeDueCheck),
            "include_completed": include_completed,
            "include_uncompleted": include_uncompleted,
            "include_category": self._is_checked(self.includeCategoryCheck),
            "category_filter": _safe_text(self.filterCategoryCombo.SelectedItem),
            "assigned_filter": _safe_text(self.filterAssignedCombo.SelectedItem),
            "start_date": _safe_text(self.startDateText.Text).strip(),
            "end_date": _safe_text(self.endDateText.Text).strip(),
            "sort_order": "oldest" if _safe_text(self.sortOrderCombo.SelectedItem) == "Oldest First" else "newest"
        }
        self.Close()

    def on_cancel(self, sender, args):
        self.result = None
        self.Close()


def prompt_export_settings(notes):
    xaml_path = script.get_bundle_file("export_options_ui.xaml")
    if not xaml_path or not os.path.exists(xaml_path):
        forms.alert("UI file not found: export_options_ui.xaml", warn_icon=True)
        return None

    window = ExportOptionsWindow(xaml_path, notes)
    window.ShowDialog()
    return window.result


def build_export_html(project_name, architect_name, notes, settings, project_address=""):
    grouped = {}

    category_filter = _safe_text(settings.get("category_filter", "")).strip()
    assigned_filter = _safe_text(settings.get("assigned_filter", "")).strip()
    start_dt = _parse_note_datetime(settings.get("start_date", "")) if _safe_text(settings.get("start_date", "")).strip() else None
    end_dt = _parse_note_datetime(settings.get("end_date", "")) if _safe_text(settings.get("end_date", "")).strip() else None
    if end_dt and end_dt != datetime.datetime.min:
        end_dt = end_dt + datetime.timedelta(days=1)

    filtered = []
    for note in notes:
        normalized = _normalize_note(note)
        if bool(normalized.get("deleted", False)):
            continue

        is_completed = bool(normalized.get("completed"))
        if (not settings.get("include_completed", True)) and is_completed:
            continue
        if (not settings.get("include_uncompleted", True)) and (not is_completed):
            continue

        if category_filter and category_filter != "All" and normalized.get("category") != _normalize_category(category_filter):
            continue

        assigned_to = _safe_text(normalized.get("assignedTo", "")).strip()
        if assigned_filter and assigned_filter != "All" and assigned_to != assigned_filter:
            continue

        if start_dt or end_dt:
            note_dt = _parse_note_datetime(normalized.get("timestamp", ""))
            if start_dt and start_dt != datetime.datetime.min and note_dt < start_dt:
                continue
            if end_dt and end_dt != datetime.datetime.min and note_dt >= end_dt:
                continue

        filtered.append(normalized)

    sort_order = settings.get("sort_order", "newest")
    reverse = (sort_order != "oldest")
    filtered = sorted(filtered, key=lambda n: _parse_note_datetime(n.get("timestamp", "")), reverse=reverse)

    for normalized in filtered:
        room_label = normalized.get("roomDisplay", "Unknown Room")
        grouped.setdefault(room_label, []).append(normalized)

    rooms_html = []
    for room_label in sorted(grouped.keys()):
        items_html = []
        room_notes = grouped[room_label]

        for note in room_notes:
            checked = bool(note.get("completed", False))
            css_class = "completed" if checked else "pending"
            badge_class = "badge-%s" % note.get("category", "Observation").replace(" ", "-").lower()
            status = "Completed" if checked else "Open"

            ts = html_escape(note.get("timestamp", ""))
            txt = html_escape(note.get("text", ""))
            category = html_escape(note.get("category", "Observation"))
            assigned_to = html_escape(note.get("assignedTo", "")) or "Unassigned"
            due_date = html_escape(note.get("dueDate", "")) or "Not set"
            comments = note.get("comments", []) if isinstance(note.get("comments", []), list) else []

            top_line = '<div class="note-top"><span class="time">%s</span></div>' % ts
            if settings.get("include_category"):
                top_line = '<div class="note-top"><span class="badge %s">%s</span><span class="time">%s</span></div>' % (badge_class, category, ts)

            meta_chunks = ['<span>Status: %s</span>' % status]
            if settings.get("include_assigned"):
                meta_chunks.insert(0, '<span>Assigned: %s</span>' % assigned_to)
            if settings.get("include_due"):
                meta_chunks.append('<span>Due: %s</span>' % due_date)

            comments_sorted = sorted(
                comments,
                key=lambda c: _parse_note_datetime(_safe_text(c.get("timestamp", ""))),
                reverse=True
            )
            comment_items = []
            for c in comments_sorted:
                c_ts = html_escape(c.get("timestamp", ""))
                c_author = html_escape(c.get("author", ""))
                c_txt = html_escape(c.get("text", ""))
                if c_author:
                    c_head = "%s | %s" % (c_ts, c_author)
                else:
                    c_head = c_ts
                comment_items.append('<li><span class="comment-meta">%s</span><div class="comment-text">%s</div></li>' % (c_head, c_txt))

            comments_html = ""
            if comment_items:
                comments_html = '<details class="comments-wrap"><summary>Show Comments (%s)</summary><ul class="comment-list">%s</ul></details>' % (
                    len(comment_items),
                    ''.join(comment_items)
                )

            items_html.append(
                '<article class="note-card %s">%s<div class="note-text">%s</div><div class="meta-row">%s</div>%s</article>'
                % (css_class, top_line, txt, "".join(meta_chunks), comments_html)
            )

        room_title = html_escape(room_label)
        if not settings.get("include_room_numbers") and room_notes:
            room_title = html_escape(room_notes[0].get("roomName", room_label) or room_label)

        rooms_html.append(
            '<section class="room">\n'
            '  <h2>%s</h2>\n'
            '  <div class="room-grid">\n%s\n  </div>\n'
            '</section>' % (room_title, "\n".join(["    " + line for line in items_html]))
        )

    header_lines = []
    if settings.get("include_project"):
        header_lines.append('<div class="meta">Project: %s</div>' % html_escape(project_name))
        header_lines.append('<div class="meta">Architect/Firm: %s</div>' % html_escape(architect_name))
        if _safe_text(project_address).strip():
            header_lines.append('<div class="meta">Address: %s</div>' % html_escape(project_address))
    if settings.get("include_datetime"):
        header_lines.append('<div class="meta">Exported: %s</div>' % html_escape(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

    html = """<!doctype html>
<html>
<head>
    <meta charset=\"utf-8\" />
    <title>Redline - %s</title>
    <style>
        body { font-family: Segoe UI, Arial, sans-serif; color: #111827; margin: 28px; background: #F9FAFB; }
        .header { margin-bottom: 20px; border-bottom: 2px solid #E5E7EB; padding-bottom: 12px; }
        .wordmark { color: #D97706; font-weight: 800; letter-spacing: 1px; font-size: 18px; }
        h1 { margin: 6px 0 4px 0; font-size: 24px; }
        .meta { color: #4B5563; margin: 2px 0; }
        .room { margin-bottom: 18px; break-inside: avoid; }
        .room h2 { font-size: 16px; margin: 0 0 10px 0; }
        .room-grid { display: block; }
        .note-card { background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 8px; padding: 10px; margin-bottom: 8px; }
        .note-top { margin-bottom: 6px; }
        .badge { display: inline-block; font-size: 11px; font-weight: 700; border-radius: 999px; padding: 3px 8px; color: #FFFFFF; margin-right: 8px; }
        .badge-decision { background: #2563EB; }
        .badge-action-item { background: #DC2626; }
        .badge-question { background: #F59E0B; color: #111827; }
        .badge-observation { background: #6B7280; }
        .time { color: #6B7280; font-size: 11px; }
        .note-text { margin: 6px 0; line-height: 1.45; }
        .meta-row { color: #4B5563; font-size: 12px; }
        .meta-row span { margin-right: 14px; }
        .completed { opacity: 0.75; }
        .comments-wrap { margin-top: 8px; }
        .comments-wrap summary { cursor: pointer; color: #1F2937; font-size: 12px; }
        .comment-list { margin: 6px 0 0 0; padding-left: 18px; }
        .comment-list li { margin: 0 0 6px 0; }
        .comment-meta { color: #6B7280; font-size: 11px; display: block; }
        .comment-text { color: #374151; font-size: 12px; }
    </style>
</head>
<body>
    <div class=\"header\">
        <div class=\"wordmark\">REDLINE</div>
        <h1>Datum Notes &mdash; Redline Export</h1>
        %s
    </div>
    %s
</body>
</html>
""" % (
        html_escape(project_name),
        "".join(header_lines),
        "\n".join(rooms_html) if rooms_html else "<p>No redlines found.</p>"
    )
    return html


def build_weekly_digest_html(project_name, notes):
    grouped = {}
    now = datetime.datetime.now()
    cutoff = now - datetime.timedelta(days=7)

    for note in notes:
        n = _normalize_note(note)
        if bool(n.get("deleted", False)):
            continue
        dt = _parse_note_datetime(n.get("timestamp", ""))
        if dt == datetime.datetime.min or dt < cutoff:
            continue
        room = _safe_text(n.get("roomDisplay", "Unknown Room")) or "Unknown Room"
        grouped.setdefault(room, []).append(n)

    rows = []
    for room in sorted(grouped.keys()):
        rows.append("<section class=\"room\"><h2>%s</h2><ul>" % html_escape(room))
        for n in sorted(grouped[room], key=lambda x: _parse_note_datetime(x.get("timestamp", "")), reverse=True):
            rows.append("<li><strong>%s</strong> | %s | %s</li>" % (
                html_escape(n.get("category", "Observation")),
                html_escape(n.get("text", "")),
                html_escape(n.get("timestamp", ""))
            ))
        rows.append("</ul></section>")

    html = """<!doctype html>
<html>
<head>
<meta charset=\"utf-8\" />
<title>Redline Weekly Digest - %s</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 28px; color: #111827; }
h1 { margin: 0 0 10px 0; }
.room { margin: 0 0 14px 0; }
.room h2 { margin: 0 0 6px 0; font-size: 16px; }
li { margin: 0 0 4px 0; }
</style>
</head>
<body>
<h1>Redline Weekly Digest</h1>
<div>Project: %s</div>
<div>Generated: %s</div>
%s
</body>
</html>
""" % (
        html_escape(project_name),
        html_escape(project_name),
        html_escape(now.strftime("%Y-%m-%d %H:%M:%S")),
        "\n".join(rows) if rows else "<p>No redlines added in the last 7 days.</p>"
    )
    return html


def build_ai_template(project_name, rooms):
    date_text = datetime.datetime.now().strftime("%Y-%m-%d")
    
    # Build room list
    room_list = ""
    if rooms:
        room_list = "AVAILABLE ROOMS: " + ", ".join(rooms[:10]) + ("\n" if len(rooms) > 10 else "")
    
    return """STRICT OUTPUT FORMAT - FOLLOW EXACTLY
======================================

Parse the transcript below and extract ALL items. Output ONLY the raw data content below - no other text, no explanations.

CRITICAL: OUTPUT THE ENTIRE RESPONSE AS ONE SINGLE LINE.
CRITICAL: NEVER WRAP TO MULTIPLE LINES.

CATEGORY MARKERS (must prefix each item):
@@ACTION - For action items that need to be completed
@@DECISION - For decisions that were made
@@QUESTION - For open questions/discussions  
@@OBSERVATION - For general notes/remarks
@@UNASSIGNED - For items where you cannot determine the room

PROJECT: %s
DATE: %s
%s

CRITICAL RULES (FOLLOW PRECISELY):
1. Entire output must be one line only (single-line response)
2. If text is long, keep it on that same line with spaces
3. Use pipe character | to separate fields (pipes ONLY between fields, not elsewhere)
4. Do NOT include the marker text (@@ACTION, @@DECISION, etc.) inside the note text
5. Fields for each type:
   - @@ACTION: room | note text | assigned to | due date (YYYY-MM-DD or leave empty)
   - @@DECISION: room | note text
   - @@QUESTION: room | note text | assigned to (or leave empty)
   - @@OBSERVATION: room | note text
   - @@UNASSIGNED: note text | category keyword (ACTION/DECISION/QUESTION/OBSERVATION)
6. For empty fields, use "" or just leave it blank
7. Room names must match the available rooms list above
8. Do NOT add line numbers, bullets, explanations, or extra formatting

EXAMPLE OUTPUT (single-line full response):
@@ACTION: Electrical Room | Replace all south wall fixtures that need attention | Electrician | 2026-03-25 @@ACTION: HVAC Zone 3 | Pressurize test report and document results | MEP Engineer | 2026-03-22 @@DECISION: Structural | Approved 8-inch reinforcement depth change @@DECISION: Architectural | Finalize paint schedule next week during coordination @@QUESTION: Parking Level B | What's the new capacity requirement in current plan? | PM @@QUESTION: Safety | Fire rating on new sealant material? |  @@OBSERVATION: Lobby | Ceiling height increased by 2 feet from original design @@OBSERVATION: Conference Room | Lighting fixtures are different from specs submitted @@UNASSIGNED: Schedule progress meeting on Friday afternoon | ACTION @@UNASSIGNED: Document all RFIs received this month | QUESTION

NOW OUTPUT ONLY THE ITEMS (single-line response, use markers above):
""" % (project_name, date_text, room_list)


def parse_ai_template_input(text):
    """Parse AI template output line-by-line using @@ category markers.

    Expected line format:
    @@CATEGORY: field1 | field2 | field3 | ...
    """
    raw = _safe_text(text)
    if not raw:
        return []

    # Normalize newlines, then split chained @@ items into separate lines.
    raw = raw.replace("\r\n", "\n").replace("\r", "\n")
    raw = re.sub(r"(?<!\n)@@", "\n@@", raw)

    parsed = []

    # Remove Markdown code-fence lines such as ``` or ```text
    lines = []
    for line in raw.split("\n"):
        cleaned = _safe_text(line).strip()
        if cleaned.startswith("```"):
            continue
        lines.append(cleaned)

    for line in lines:
        if not line:
            continue

        upper_line = line.upper()
        marker = ""
        content = ""

        if upper_line.startswith("@@ACTION"):
            marker = "ACTION"
            content = line[len("@@ACTION"):].lstrip(": ").strip()
        elif upper_line.startswith("@@DECISION"):
            marker = "DECISION"
            content = line[len("@@DECISION"):].lstrip(": ").strip()
        elif upper_line.startswith("@@QUESTION"):
            marker = "QUESTION"
            content = line[len("@@QUESTION"):].lstrip(": ").strip()
        elif upper_line.startswith("@@OBSERVATION"):
            marker = "OBSERVATION"
            content = line[len("@@OBSERVATION"):].lstrip(": ").strip()
        elif upper_line.startswith("@@UNASSIGNED"):
            marker = "UNASSIGNED"
            content = line[len("@@UNASSIGNED"):].lstrip(": ").strip()
        else:
            # Ignore lines that do not start with a recognized marker.
            continue

        if not content:
            continue

        parts = [p.strip() for p in content.split("|")]
        parts = [p if p != '""' and p != "''" else "" for p in parts]

        assigned = ""
        due = ""
        is_unassigned = False

        if marker == "ACTION":
            if len(parts) < 2:
                continue
            room = parts[0] or "UNASSIGNED"
            note = parts[1]
            category = "Action Item"
            assigned = parts[2] if len(parts) > 2 else ""
            due = parts[3] if len(parts) > 3 else ""
        elif marker == "DECISION":
            if len(parts) < 2:
                continue
            room = parts[0] or "UNASSIGNED"
            note = parts[1]
            category = "Decision"
        elif marker == "QUESTION":
            if len(parts) < 2:
                continue
            room = parts[0] or "UNASSIGNED"
            note = parts[1]
            category = "Question"
            assigned = parts[2] if len(parts) > 2 else ""
        elif marker == "OBSERVATION":
            if len(parts) < 2:
                continue
            room = parts[0] or "UNASSIGNED"
            note = parts[1]
            category = "Observation"
        else:
            if len(parts) < 1:
                continue
            room = "UNASSIGNED"
            note = parts[0]
            category_hint = parts[1] if len(parts) > 1 else "Observation"
            category = _normalize_category(category_hint)
            is_unassigned = True

        if _safe_text(note).strip():
            parsed.append({
                "room": room,
                "text": note,
                "category": category,
                "assignedTo": assigned,
                "dueDate": due,
                "isUnassigned": is_unassigned
            })

    return parsed


class RedlineWindow(forms.WPFWindow):
    def __init__(self, xaml_path, doc):
        forms.WPFWindow.__init__(self, xaml_path)

        self.doc = doc
        self.current_user = get_current_user_name(doc)
        self.rooms = collect_rooms(doc)
        self.room_lookup = {}
        self.selected_room = None
        self.active_tab = "All"
        self.show_completed = False
        self.show_deleted = False
        self.room_expanders = []
        self._room_resolved_cache = {}
        self.selected_work_day = ""

        self.config = load_redline_config()
        self.manual_resolved_rooms = set([_safe_text(x) for x in self.config.get("manualResolvedRooms", []) if _safe_text(x)])
        self.last_custom_assignee = _safe_text(self.config.get("lastCustomAssignee", "")).strip()

        self.store_path = json_path_for_document(doc)
        self.all_notes = load_notes(self.store_path)
        self.work_log_items = load_work_log_items(self.store_path)

        _project_data = load_project_payload(self.store_path)
        _saved_custom = _project_data.get("customAssignees", []) if isinstance(_project_data, dict) else []
        self.custom_assignees = sorted(list(set([_safe_text(x).strip() for x in _saved_custom if _safe_text(x).strip()])))

        if self._ensure_note_ids():
            self._save()

        if self._migrate_assignee_fields():
            self._save()

        if self._purge_deleted_notes():
            self._save()

        self._wire_events()
        self.versionTextRun.Text = "v%s" % EXTENSION_VERSION
        self._bind_rooms()
        self._bind_categories()
        self._bind_assignees()
        self._bind_history_sort()
        self._bind_history_filter()
        self._update_tab_visuals()
        self._update_toggle_button_text()
        self._toggle_due_date_visibility()
        self._update_selected_room_ui()
        self._render_work_history()
        self._render_history()

    def _migrate_assignee_fields(self):
        """Unify assignee keys so stale legacy values do not reappear in UI."""
        changed = False
        migrated = []

        for note in self.all_notes:
            n = note if isinstance(note, dict) else {}

            val_assigned_to = _safe_text(n.get("assigned_to", "")).strip()
            val_assignedTo = _safe_text(n.get("assignedTo", "")).strip()
            val_assigned = _safe_text(n.get("assigned", "")).strip()
            val_assignee = _safe_text(n.get("assignee", "")).strip()

            # Pick best non-empty value with preference to assigned_to then assignedTo.
            merged = val_assigned_to or val_assignedTo or val_assigned or val_assignee

            if _safe_text(n.get("assignedTo", "")).strip() != merged:
                n["assignedTo"] = merged
                changed = True
            if _safe_text(n.get("assigned_to", "")).strip() != merged:
                n["assigned_to"] = merged
                changed = True

            # Remove extra legacy aliases to prevent future conflicts.
            if "assigned" in n:
                try:
                    del n["assigned"]
                    changed = True
                except Exception:
                    pass
            if "assignee" in n:
                try:
                    del n["assignee"]
                    changed = True
                except Exception:
                    pass

            migrated.append(n)

        if changed:
            self.all_notes = migrated

        return changed

    def _ensure_note_ids(self):
        """Guarantee each note has a unique stable id so card actions work correctly.
        Fixes both missing IDs and duplicate IDs (which arise when notes were bulk-stamped
        in the same microsecond on a first run with no existing IDs)."""
        changed = False
        seen_ids = set()
        for i, note in enumerate(self.all_notes):
            raw = note if isinstance(note, dict) else {}
            existing = _safe_text(raw.get("id", "")).strip()
            # Assign a new ID if missing OR if this ID was already seen (duplicate).
            if not existing or existing in seen_ids:
                raw["id"] = "N%s_%d" % (datetime.datetime.now().strftime("%Y%m%d%H%M%S%f"), i)
                self.all_notes[i] = raw
                changed = True
                seen_ids.add(raw["id"])
            else:
                seen_ids.add(existing)
        return changed

    def _save_config(self):
        self.config["manualResolvedRooms"] = sorted(list(self.manual_resolved_rooms))
        self.config["lastCustomAssignee"] = _safe_text(self.last_custom_assignee).strip()
        save_redline_config(self.config)

    def _remember_custom_assignee(self, value):
        clean = _safe_text(value).strip()
        if not clean:
            return

        changed = False
        if clean not in self.custom_assignees:
            self.custom_assignees.append(clean)
            self.custom_assignees = sorted(list(set(self.custom_assignees)))
            changed = True

        if clean != self.last_custom_assignee:
            self.last_custom_assignee = clean
            changed = True

        if changed:
            self._save_config()

    def _wire_events(self):
        self.tutorialButton.Click += self.on_open_tutorial
        self.addNoteButton.Click += self.on_add_note
        self.addWorkItemButton.Click += self.on_add_work_item
        self.updateButton.Click += self.on_open_update_info
        self.calendarWorkButton.Click += self.on_pick_work_day
        self.exportButton.Click += self.on_export
        self.bulkAssignButton.Click += self.on_bulk_assign
        self.bulkTypeButton.Click += self.on_bulk_type
        self.weeklyDigestButton.Click += self.on_weekly_digest_export
        self.exportTeamButton.Click += self.on_export_for_team_member
        self.importTeamButton.Click += self.on_import_team_file
        self.copyAiTemplateButton.Click += self.on_copy_ai_template
        self.importAiButton.Click += self.on_import_from_ai
        self.uploadLinkButton.Click += self.on_open_upload_info

        self.categoryCombo.SelectionChanged += self.on_category_changed
        self.assignedToCombo.SelectionChanged += self.on_assigned_to_changed
        self.addCustomAssigneeButton.Click += self.on_add_custom_assignee

        self.historyFilterCombo.SelectionChanged += self.on_filter_changed
        self.historySortCombo.SelectionChanged += self.on_sort_changed
        self.historySearchBox.TextChanged += self.on_search_changed

        self.roomCombo.SelectionChanged += self.on_room_changed
        self.goToRoomButton.Click += self.on_go_to_room
        self.toggleRoomResolvedButton.Click += self.on_toggle_room_resolved

        self.expandAllButton.Click += self.on_expand_all
        self.collapseAllButton.Click += self.on_collapse_all
        self.tabAllButton.Click += self.on_tab_all
        self.tabActionButton.Click += self.on_tab_action
        self.tabQuestionButton.Click += self.on_tab_question
        self.tabDecisionButton.Click += self.on_tab_decision
        self.tabObservationButton.Click += self.on_tab_observation
        self.tabUnassignedButton.Click += self.on_tab_unassigned
        self.toggleCompletedButton.Click += self.on_toggle_completed
        self.toggleDeletedButton.Click += self.on_toggle_deleted

        self.noteText.KeyDown += self.on_note_text_keydown
        self.workTodayText.KeyDown += self.on_work_text_keydown
        self.PreviewKeyDown += self.on_window_keydown

    def _purge_deleted_notes(self):
        now = datetime.datetime.now()
        cutoff = now - datetime.timedelta(days=7)
        changed = False
        kept = []

        for note in self.all_notes:
            n = _normalize_note(note)
            if bool(n.get("deleted", False)):
                deleted_dt = _parse_note_datetime(n.get("deletedAt", ""))
                if deleted_dt == datetime.datetime.min or deleted_dt < cutoff:
                    changed = True
                    continue
            kept.append(n)

        if changed:
            self.all_notes = kept

        return changed

    def _active_notes(self):
        active = []
        for note in self.all_notes:
            raw = note if isinstance(note, dict) else {}
            if not bool(raw.get("deleted", False)):
                active.append(note)
        return active

    def _deleted_notes(self):
        deleted = []
        for note in self.all_notes:
            raw = note if isinstance(note, dict) else {}
            if bool(raw.get("deleted", False)):
                deleted.append(note)
        return deleted

    def _room_is_resolved(self, room):
        if not room:
            return False

        room_id = _safe_text(room.get("roomId"))
        if room_id in self._room_resolved_cache:
            return bool(self._room_resolved_cache.get(room_id))

        if room_id in self.manual_resolved_rooms:
            return True

        room_notes = []
        for note in self.all_notes:
            n = _normalize_note(note)
            if bool(n.get("deleted", False)):
                continue
            if _safe_text(n.get("roomId")) == room_id:
                room_notes.append(n)

        action_items = [n for n in room_notes if _normalize_category(n.get("category")) == "Action Item"]
        if not action_items:
            return False

        for item in action_items:
            if not bool(item.get("completed", False)):
                return False
        return True

    def _rebuild_room_resolved_cache(self, normalized_active=None):
        notes = normalized_active if normalized_active is not None else [_normalize_note(n) for n in self._active_notes()]

        action_totals = {}
        action_open = {}
        for n in notes:
            rid = _safe_text(n.get("roomId", ""))
            if not rid:
                continue
            if _normalize_category(n.get("category")) != "Action Item":
                continue

            action_totals[rid] = int(action_totals.get(rid, 0)) + 1
            if not bool(n.get("completed", False)):
                action_open[rid] = int(action_open.get(rid, 0)) + 1

        cache = {}
        for r in self.rooms:
            rid = _safe_text(r.get("roomId", ""))
            if not rid:
                continue
            if rid in self.manual_resolved_rooms:
                cache[rid] = True
            else:
                total = int(action_totals.get(rid, 0))
                cache[rid] = bool(total > 0 and int(action_open.get(rid, 0)) == 0)

        self._room_resolved_cache = cache

    def _room_label(self, room):
        base = _safe_text(room.get("roomDisplay", "Unknown Room"))
        if self._room_is_resolved(room):
            return "[CHECK] %s" % base
        return base

    def _bind_rooms(self):
        self.roomCombo.Items.Clear()
        self.room_lookup = {}

        general_room = unassigned_room_bucket()
        general_label = _safe_text(general_room.get("roomDisplay", ""))
        self.room_lookup[general_label] = general_room
        self.roomCombo.Items.Add(general_label)

        for room in self.rooms:
            label = self._room_label(room)
            self.room_lookup[label] = room
            self.roomCombo.Items.Add(label)

        self.roomCombo.IsEnabled = True
        self.roomCombo.SelectedIndex = 0
        self.selected_room = self.room_lookup.get(_safe_text(self.roomCombo.SelectedItem))

    def _bind_categories(self):
        self.categoryCombo.Items.Clear()
        for category in CATEGORY_OPTIONS:
            self.categoryCombo.Items.Add(category)
        self.categoryCombo.SelectedItem = "Observation"

    def _bind_assignees(self):
        prev = _safe_text(getattr(self.assignedToCombo, "SelectedItem", ""))
        self.assignedToCombo.Items.Clear()

        for name in DEFAULT_TEAM_MEMBERS:
            self.assignedToCombo.Items.Add(name)

        if self.custom_assignees:
            self.assignedToCombo.Items.Add("--- Saved Custom Names ---")
            for name in self.custom_assignees:
                self.assignedToCombo.Items.Add("Custom: %s" % name)

        self.assignedToCombo.Items.Add("Custom...")

        if prev and prev in [ _safe_text(x) for x in self.assignedToCombo.Items ]:
            self.assignedToCombo.SelectedItem = prev
        else:
            self.assignedToCombo.SelectedIndex = 0

        self._toggle_custom_assignee_ui()

    def _toggle_custom_assignee_ui(self):
        selected = _safe_text(self.assignedToCombo.SelectedItem)
        show = (selected == "Custom...")
        self.assignedToCustomText.Visibility = Visibility.Visible if show else Visibility.Collapsed
        self.addCustomAssigneeButton.Visibility = Visibility.Visible if show else Visibility.Collapsed
        if show and not _safe_text(self.assignedToCustomText.Text).strip() and self.last_custom_assignee:
            self.assignedToCustomText.Text = self.last_custom_assignee
        if not show:
            self.assignedToCustomText.Text = ""

    def _bind_history_sort(self):
        self.historySortCombo.Items.Clear()
        for item in SORT_OPTIONS:
            self.historySortCombo.Items.Add(item)
        self.historySortCombo.SelectedItem = "Newest First"

    def _bind_history_filter(self):
        previous = _safe_text(getattr(self.historyFilterCombo, "SelectedItem", ""))

        self.historyFilterCombo.Items.Clear()
        self.historyFilterCombo.Items.Add("All Rooms")

        room_labels = set()
        for note in self._active_notes():
            n = _normalize_note(note)
            room_labels.add(_safe_text(n.get("roomDisplay", "Unknown Room")) or "Unknown Room")

        for label in sorted(room_labels):
            self.historyFilterCombo.Items.Add(label)

        if previous and previous in room_labels:
            self.historyFilterCombo.SelectedItem = previous
        else:
            self.historyFilterCombo.SelectedItem = "All Rooms"

    def _toggle_due_date_visibility(self):
        selected = _safe_text(self.categoryCombo.SelectedItem)
        if selected == "Action Item":
            self.dueDatePanel.Visibility = Visibility.Visible
        else:
            self.dueDatePanel.Visibility = Visibility.Collapsed
            self.dueDatePicker.SelectedDate = None

    def _filtered_notes(self, include_completed=True, include_deleted=False, pre_normalized=None):
        selected_room = _safe_text(self.historyFilterCombo.SelectedItem)
        search_text = _safe_text(self.historySearchBox.Text).strip().lower()

        if pre_normalized is not None:
            notes = list(pre_normalized)
        elif include_deleted:
            notes = [_normalize_note(n) for n in self._deleted_notes()]
        else:
            notes = [_normalize_note(n) for n in self._active_notes()]

        if selected_room and selected_room != "All Rooms":
            notes = [n for n in notes if (_safe_text(n.get("roomDisplay", "Unknown Room")) or "Unknown Room") == selected_room]

        if self.active_tab != "All":
            if self.active_tab == "Pending":
                notes = [n for n in notes if bool(n.get("pending", False)) and not bool(n.get("completed", False))]
            else:
                notes = [n for n in notes if _normalize_category(n.get("category")) == self.active_tab]

        if not include_completed:
            notes = [n for n in notes if not bool(n.get("completed", False))]

        if search_text:
            def _matches(n):
                body = " ".join([
                    _safe_text(n.get("text", "")),
                    _safe_text(n.get("assignedTo", "")),
                    _safe_text(n.get("roomDisplay", "")),
                    _safe_text(n.get("category", "")),
                    _safe_text(n.get("duplicateFrom", ""))
                ]).lower()
                return search_text in body
            notes = [n for n in notes if _matches(n)]

        sort_name = _safe_text(self.historySortCombo.SelectedItem) or "Newest First"
        if sort_name == "By Room":
            notes = sorted(notes, key=lambda n: _parse_note_datetime(n.get("timestamp", "")), reverse=True)
            notes = sorted(notes, key=lambda n: (_safe_text(n.get("roomDisplay", "Unknown Room")) or "Unknown Room"))
        elif sort_name == "Oldest First":
            notes = sorted(notes, key=lambda n: _parse_note_datetime(n.get("timestamp", "")))
        else:
            notes = sorted(notes, key=lambda n: _parse_note_datetime(n.get("timestamp", "")), reverse=True)

        return notes

    def _group_notes_by_room(self, notes):
        grouped = {}
        for note in notes:
            room = _safe_text(note.get("roomDisplay", "Unknown Room")) or "Unknown Room"
            grouped.setdefault(room, []).append(note)
        return grouped

    def _set_active_tab(self, tab_name):
        self.active_tab = tab_name if tab_name in CATEGORY_TAB_OPTIONS else "All"
        self._update_tab_visuals()
        try:
            self._render_history()
        except Exception as ex:
            forms.alert(
                "Could not refresh this tab.\n%s\n\n%s" % (_safe_text(ex), traceback.format_exc()),
                warn_icon=True
            )

    def _set_tab_button_style(self, button, is_active):
        button.Background = _brush("#F59E0B") if is_active else _brush(THEME_BUTTON_BG)
        button.Foreground = _brush("#111111") if is_active else _brush("#F3F4F6")
        button.BorderBrush = _brush("#D97706") if is_active else _brush(THEME_BUTTON_BORDER)

    def _update_tab_visuals(self):
        self._set_tab_button_style(self.tabAllButton, self.active_tab == "All")
        self._set_tab_button_style(self.tabActionButton, self.active_tab == "Action Item")
        self._set_tab_button_style(self.tabQuestionButton, self.active_tab == "Question")
        self._set_tab_button_style(self.tabDecisionButton, self.active_tab == "Decision")
        self._set_tab_button_style(self.tabObservationButton, self.active_tab == "Observation")
        self._set_tab_button_style(self.tabUnassignedButton, self.active_tab == "Pending")

    def _update_toggle_button_text(self):
        self.toggleCompletedButton.Content = "Hide Completed" if self.show_completed else "Show Completed"
        self.toggleDeletedButton.Content = "Hide Deleted" if self.show_deleted else "Show Deleted"

    def _room_stats(self, room):
        if not room:
            return (0, 0, 0)

        room_id = _safe_text(room.get("roomId"))
        notes = [n for n in self._active_notes() if _safe_text(_normalize_note(n).get("roomId")) == room_id]
        notes = [_normalize_note(n) for n in notes]
        total = len(notes)
        completed = len([n for n in notes if bool(n.get("completed", False))])
        open_actions = len([n for n in notes if _normalize_category(n.get("category")) == "Action Item" and not bool(n.get("completed", False))])
        return (total, completed, open_actions)

    def _global_stats(self):
        notes = [_normalize_note(n) for n in self._active_notes()]
        total = len(notes)
        action_items = len([n for n in notes if _normalize_category(n.get("category")) == "Action Item"])
        completed = len([n for n in notes if bool(n.get("completed", False))])
        return (total, action_items, completed)

    def _update_selected_room_ui(self):
        room = self.selected_room
        if not room:
            self.roomPrimaryText.Text = "Room: -"
            self.roomSecondaryText.Text = "Level: -"
            self.roomSummaryText.Text = "Redline summary: -"
            self.toggleRoomResolvedButton.Content = "Mark Room Resolved"
            return

        room_number = _safe_text(room.get("number")) or "?"
        room_name = _safe_text(room.get("name")) or "Unnamed"
        level = _safe_text(room.get("level")) or "Unknown"
        total, completed, open_actions = self._room_stats(room)
        resolved = self._room_is_resolved(room)

        self.roomPrimaryText.Text = "Room %s - %s%s" % (room_number, room_name, " [CHECK]" if resolved else "")
        self.roomSecondaryText.Text = "Level: %s" % level
        self.roomSummaryText.Text = "Redlines: %s | Completed: %s | Open Action Items: %s" % (total, completed, open_actions)
        self.toggleRoomResolvedButton.Content = "Unmark Room Resolved" if _safe_text(room.get("roomId")) in self.manual_resolved_rooms else "Mark Room Resolved"

    def _build_comments_stack(self, note):
        comments = note.get("comments", []) if isinstance(note.get("comments", []), list) else []

        comments_stack = StackPanel()

        add_row = Grid()
        add_row.ColumnDefinitions.Add(ColumnDefinition())
        add_row.ColumnDefinitions.Add(ColumnDefinition())
        add_row.ColumnDefinitions[1].Width = GridLength.Auto

        comment_input = TextBox()
        comment_input.MinHeight = 28
        comment_input.Margin = Thickness(0, 0, 8, 0)
        comment_input.Background = _brush(THEME_INPUT_BG)
        comment_input.Foreground = _brush("#F3F4F6")
        comment_input.BorderBrush = _brush(THEME_BUTTON_BORDER)
        comment_input.Tag = note.get("id")
        Grid.SetColumn(comment_input, 0)
        add_row.Children.Add(comment_input)

        add_comment_btn = Button()
        add_comment_btn.Content = "Add Comment"
        add_comment_btn.Tag = comment_input
        add_comment_btn.Padding = Thickness(8, 3, 8, 3)
        add_comment_btn.Background = _brush(THEME_BUTTON_BG)
        add_comment_btn.Foreground = _brush("#F3F4F6")
        add_comment_btn.BorderBrush = _brush(THEME_BUTTON_BORDER)
        add_comment_btn.Click += self.on_add_comment
        Grid.SetColumn(add_comment_btn, 1)
        add_row.Children.Add(add_comment_btn)

        comments_stack.Children.Add(add_row)

        comments_list = sorted(
            comments,
            key=lambda c: _parse_note_datetime(_safe_text(c.get("timestamp", ""))),
            reverse=True
        )
        for c in comments_list:
            c_ts = _safe_text(c.get("timestamp", ""))
            c_author = _safe_text(c.get("author", ""))
            c_txt = _safe_text(c.get("text", ""))

            item_border = Border()
            item_border.Margin = Thickness(0, 6, 0, 0)
            item_border.Padding = Thickness(6)
            item_border.Background = _brush(THEME_PANEL_BG)
            item_border.BorderBrush = _brush(THEME_PANEL_BORDER)
            item_border.BorderThickness = Thickness(1)

            item_stack = StackPanel()

            item_meta = TextBlock()
            if c_author:
                item_meta.Text = "%s | %s" % (c_ts, c_author)
            else:
                item_meta.Text = c_ts
            item_meta.FontSize = 10
            item_meta.Foreground = _brush("#9CA3AF")
            item_stack.Children.Add(item_meta)

            item_text = TextBlock()
            item_text.Text = c_txt
            item_text.TextWrapping = TextWrapping.Wrap
            item_text.Foreground = _brush("#E5E7EB")
            item_text.Margin = Thickness(0, 2, 0, 0)
            item_stack.Children.Add(item_text)

            item_border.Child = item_stack
            comments_stack.Children.Add(item_border)

        return comments_stack

    def on_comments_expander_expanded(self, sender, args):
        payload = sender.Tag if isinstance(sender.Tag, dict) else {}
        if bool(payload.get("loaded", False)):
            return

        note = payload.get("note") if isinstance(payload, dict) else None
        note = note if isinstance(note, dict) else {}
        sender.Content = self._build_comments_stack(note)
        payload["loaded"] = True
        sender.Tag = payload

    def _make_note_card(self, note, deleted_mode=False):
        category = _normalize_category(note.get("category", "Observation"))
        comments_count = len(note.get("comments", [])) if isinstance(note.get("comments", []), list) else 0

        card = Border()
        card.Margin = Thickness(2, 0, 2, 8)
        card.Padding = Thickness(10)
        card.Background = _brush(THEME_CARD_BG)
        card.BorderBrush = _brush(BADGE_COLORS.get(category, "#6B7280"))
        card.BorderThickness = Thickness(3, 0, 0, 0)

        content = StackPanel()

        top_grid = Grid()
        top_grid.ColumnDefinitions.Add(ColumnDefinition())
        top_grid.ColumnDefinitions.Add(ColumnDefinition())
        top_grid.ColumnDefinitions[1].Width = GridLength.Auto

        top_left = StackPanel()
        top_left.Orientation = Orientation.Horizontal

        badge = TextBlock()
        badge.Text = " %s " % category
        badge.Foreground = _brush("#111111") if category == "Question" else _brush("#FFFFFF")
        badge.Background = _brush(BADGE_COLORS.get(category, "#6B7280"))
        badge.Margin = Thickness(0, 0, 8, 0)
        top_left.Children.Add(badge)

        if bool(note.get("imported", False)):
            imported = TextBlock()
            imported.Text = " Imported "
            imported.Background = _brush("#065F46")
            imported.Foreground = _brush("#ECFDF5")
            imported.Margin = Thickness(0, 0, 8, 0)
            top_left.Children.Add(imported)

        if _safe_text(note.get("duplicateFrom", "")):
            dup_badge = TextBlock()
            dup_badge.Text = " Duplicated "
            dup_badge.Background = _brush("#1D4ED8")
            dup_badge.Foreground = _brush("#DBEAFE")
            dup_badge.Margin = Thickness(0, 0, 8, 0)
            top_left.Children.Add(dup_badge)

        comment_badge = TextBlock()
        comment_badge.Text = " C:%s " % comments_count
        comment_badge.Background = _brush("#374151")
        comment_badge.Foreground = _brush("#E5E7EB")
        comment_badge.Margin = Thickness(0, 0, 8, 0)
        top_left.Children.Add(comment_badge)

        ts = TextBlock()
        ts.Text = _safe_text(note.get("timestamp", ""))
        ts.Foreground = _brush("#A3A3A3")
        ts.FontSize = 11
        ts.VerticalAlignment = VerticalAlignment.Center
        top_left.Children.Add(ts)

        edited_at = _safe_text(note.get("editedAt", "")).strip()
        if edited_at:
            edited_tag = TextBlock()
            edited_tag.Text = "  edited %s" % edited_at
            edited_tag.Foreground = _brush("#9CA3AF")
            edited_tag.FontSize = 10
            edited_tag.VerticalAlignment = VerticalAlignment.Center
            top_left.Children.Add(edited_tag)

        Grid.SetColumn(top_left, 0)
        top_grid.Children.Add(top_left)

        if not deleted_mode:
            action_wrap = StackPanel()
            action_wrap.Orientation = Orientation.Horizontal

            room_btn = Button()
            room_btn.Content = "Room"
            room_btn.Width = 48
            room_btn.Height = 22
            room_btn.Padding = Thickness(0)
            room_btn.Margin = Thickness(0, 0, 6, 0)
            room_btn.HorizontalAlignment = HorizontalAlignment.Right
            room_btn.Tag = note.get("id")
            room_btn.Background = _brush("#374151")
            room_btn.Foreground = _brush("#E5E7EB")
            room_btn.BorderBrush = _brush("#4B5563")
            room_btn.Click += self.on_reassign_note_room
            action_wrap.Children.Add(room_btn)

            who_btn = Button()
            who_btn.Content = "Who"
            who_btn.Width = 40
            who_btn.Height = 22
            who_btn.Padding = Thickness(0)
            who_btn.Margin = Thickness(0, 0, 6, 0)
            who_btn.HorizontalAlignment = HorizontalAlignment.Right
            who_btn.Tag = note.get("id")
            who_btn.Background = _brush("#334155")
            who_btn.Foreground = _brush("#E5E7EB")
            who_btn.BorderBrush = _brush("#475569")
            who_btn.Click += self.on_reassign_note_assignee
            action_wrap.Children.Add(who_btn)

            type_btn = Button()
            type_btn.Content = "Type"
            type_btn.Width = 44
            type_btn.Height = 22
            type_btn.Padding = Thickness(0)
            type_btn.Margin = Thickness(0, 0, 6, 0)
            type_btn.HorizontalAlignment = HorizontalAlignment.Right
            type_btn.Tag = note.get("id")
            type_btn.Background = _brush("#3F3F46")
            type_btn.Foreground = _brush("#F3F4F6")
            type_btn.BorderBrush = _brush("#52525B")
            type_btn.Click += self.on_reassign_note_category
            action_wrap.Children.Add(type_btn)

            edit_btn = Button()
            edit_btn.Content = "Edit"
            edit_btn.Width = 40
            edit_btn.Height = 22
            edit_btn.Padding = Thickness(0)
            edit_btn.Margin = Thickness(0, 0, 6, 0)
            edit_btn.HorizontalAlignment = HorizontalAlignment.Right
            edit_btn.Tag = note.get("id")
            edit_btn.Background = _brush("#1F2937")
            edit_btn.Foreground = _brush("#E5E7EB")
            edit_btn.BorderBrush = _brush("#374151")
            edit_btn.Click += self.on_edit_note
            action_wrap.Children.Add(edit_btn)

            dup_btn = Button()
            dup_btn.Content = "Dup"
            dup_btn.Width = 38
            dup_btn.Height = 22
            dup_btn.Padding = Thickness(0)
            dup_btn.Margin = Thickness(0, 0, 6, 0)
            dup_btn.HorizontalAlignment = HorizontalAlignment.Right
            dup_btn.Tag = note.get("id")
            dup_btn.Background = _brush("#1E3A8A")
            dup_btn.Foreground = _brush("#DBEAFE")
            dup_btn.BorderBrush = _brush("#1D4ED8")
            dup_btn.Click += self.on_duplicate_note
            action_wrap.Children.Add(dup_btn)

            delete_btn = Button()
            delete_btn.Content = "Del"
            delete_btn.Width = 36
            delete_btn.Height = 22
            delete_btn.Padding = Thickness(0)
            delete_btn.HorizontalAlignment = HorizontalAlignment.Right
            delete_btn.Tag = note.get("id")
            delete_btn.Background = _brush("#3F1D1D")
            delete_btn.Foreground = _brush("#FEE2E2")
            delete_btn.BorderBrush = _brush("#7F1D1D")
            delete_btn.Click += self.on_delete_note
            action_wrap.Children.Add(delete_btn)

            Grid.SetColumn(action_wrap, 1)
            top_grid.Children.Add(action_wrap)

        content.Children.Add(top_grid)

        body = TextBlock()
        body.Text = _safe_text(note.get("text", ""))
        body.TextWrapping = TextWrapping.Wrap
        body.Margin = Thickness(0, 8, 0, 6)
        body.Foreground = _brush("#E5E7EB")
        content.Children.Add(body)

        meta_parts = []
        assigned_to = _safe_text(note.get("assignedTo", ""))
        due_date = _safe_text(note.get("dueDate", ""))
        completed_by = _safe_text(note.get("completedBy", ""))
        completed_at = _safe_text(note.get("completedAt", ""))

        if assigned_to:
            meta_parts.append("Assigned: %s" % assigned_to)
        if due_date:
            meta_parts.append("Due: %s" % due_date)
        if completed_by:
            meta_parts.append("Completed By: %s" % completed_by)
        if completed_at:
            meta_parts.append("Completed At: %s" % completed_at)
        if _safe_text(note.get("duplicateFrom", "")):
            meta_parts.append("Duplicated from: %s" % _safe_text(note.get("duplicateFrom", "")))

        if meta_parts:
            meta = TextBlock()
            meta.Text = " | ".join(meta_parts)
            meta.FontSize = 11
            meta.Foreground = _brush("#A3A3A3")
            meta.Margin = Thickness(0, 0, 0, 6)
            content.Children.Add(meta)

        if not deleted_mode:
            comments = note.get("comments", []) if isinstance(note.get("comments", []), list) else []

            comments_expander = Expander()
            comments_expander.Margin = Thickness(0, 2, 0, 6)
            comments_expander.Foreground = _brush("#D4D4D8")
            comments_expander.IsExpanded = False
            if comments:
                comments_expander.Header = "Comments (%s)" % len(comments)
            else:
                comments_expander.Header = "Add Context"

            comments_expander.Tag = {
                "note": note,
                "loaded": False
            }
            comments_placeholder = TextBlock()
            comments_placeholder.Text = "Expand to view and add comments"
            comments_placeholder.Margin = Thickness(0, 4, 0, 4)
            comments_placeholder.Foreground = _brush("#9CA3AF")
            comments_expander.Content = comments_placeholder
            comments_expander.Expanded += self.on_comments_expander_expanded
            content.Children.Add(comments_expander)

        if deleted_mode:
            actions = StackPanel()
            actions.Orientation = Orientation.Horizontal

            restore_btn = Button()
            restore_btn.Content = "Restore"
            restore_btn.Tag = note.get("id")
            restore_btn.Margin = Thickness(0, 2, 8, 0)
            restore_btn.Padding = Thickness(8, 3, 8, 3)
            restore_btn.Background = _brush("#1F2937")
            restore_btn.Foreground = _brush("#F3F4F6")
            restore_btn.BorderBrush = _brush("#374151")
            restore_btn.Click += self.on_restore_note
            actions.Children.Add(restore_btn)

            hard_btn = Button()
            hard_btn.Content = "Delete Forever"
            hard_btn.Tag = note.get("id")
            hard_btn.Margin = Thickness(0, 2, 0, 0)
            hard_btn.Padding = Thickness(8, 3, 8, 3)
            hard_btn.Background = _brush("#7F1D1D")
            hard_btn.Foreground = _brush("#FEE2E2")
            hard_btn.BorderBrush = _brush("#991B1B")
            hard_btn.Click += self.on_hard_delete_note
            actions.Children.Add(hard_btn)

            content.Children.Add(actions)
        else:
            checkbox_row = StackPanel()
            checkbox_row.Orientation = Orientation.Horizontal

            box = CheckBox()
            box.Margin = Thickness(0, 2, 8, 0)
            box.Content = "Mark complete"
            box.IsChecked = bool(note.get("completed", False))
            box.Tag = note.get("id")
            box.Foreground = _brush("#D4D4D8")
            box.Checked += self.on_note_toggled
            box.Unchecked += self.on_note_toggled
            checkbox_row.Children.Add(box)

            pending_box = CheckBox()
            pending_box.Margin = Thickness(0, 2, 0, 0)
            pending_box.Content = "Mark pending"
            pending_box.IsChecked = bool(note.get("pending", False))
            pending_box.Tag = note.get("id")
            pending_box.Foreground = _brush("#F59E0B")
            pending_box.Checked += self.on_note_pending_toggled
            pending_box.Unchecked += self.on_note_pending_toggled
            checkbox_row.Children.Add(pending_box)

            content.Children.Add(checkbox_row)

        card.Child = content
        return card

    def on_room_expander_expanded(self, sender, args):
        payload = sender.Tag if isinstance(sender.Tag, dict) else {}
        if bool(payload.get("loaded", False)):
            return

        notes = payload.get("openNotes") if isinstance(payload.get("openNotes"), list) else []
        room_stack = StackPanel()

        if notes:
            for note in notes:
                try:
                    room_stack.Children.Add(self._make_note_card(note, deleted_mode=False))
                except Exception as ex:
                    bad = TextBlock()
                    bad.Text = "Could not render one note: %s" % _safe_text(ex)
                    bad.Margin = Thickness(6, 2, 6, 6)
                    bad.Foreground = _brush("#FCA5A5")
                    room_stack.Children.Add(bad)
        else:
            empty_room = TextBlock()
            empty_room.Text = "No open items in this room."
            empty_room.Foreground = _brush("#737373")
            empty_room.Margin = Thickness(6, 2, 6, 6)
            room_stack.Children.Add(empty_room)

        sender.Content = room_stack
        payload["loaded"] = True
        sender.Tag = payload

    def on_room_expander_collapsed(self, sender, args):
        payload = sender.Tag if isinstance(sender.Tag, dict) else {}
        if not isinstance(payload, dict):
            return

        if not bool(payload.get("loaded", False)):
            return

        placeholder = TextBlock()
        placeholder.Text = "Expand to view notes"
        placeholder.Margin = Thickness(6, 2, 6, 6)
        placeholder.Foreground = _brush("#737373")
        sender.Content = placeholder
        payload["loaded"] = False
        sender.Tag = payload

    def _render_history(self):
        self.historyPanel.Children.Clear()
        self.room_expanders = []
        if self._purge_deleted_notes():
            self._save()
            self._bind_history_filter()

        active_normalized = [_normalize_note(n) for n in self._active_notes()]
        deleted_normalized = [_normalize_note(n) for n in self._deleted_notes()]
        self._rebuild_room_resolved_cache(active_normalized)

        room_display_to_id = {}
        for r in self.rooms:
            room_display_to_id[_safe_text(r.get("roomDisplay", ""))] = _safe_text(r.get("roomId", ""))

        total = len(active_normalized)
        action_items = len([n for n in active_normalized if _normalize_category(n.get("category")) == "Action Item"])
        completed = len([n for n in active_normalized if bool(n.get("completed", False))])
        self.historyGlobalSummaryText.Text = "Total notes: %s | Action items: %s | Completed: %s" % (total, action_items, completed)
        self._update_toggle_button_text()

        all_filtered_active = self._filtered_notes(include_completed=True, include_deleted=False, pre_normalized=active_normalized)
        grouped_all = self._group_notes_by_room(all_filtered_active)

        if not grouped_all:
            empty = TextBlock()
            empty.Text = "No redlines match the current filters."
            empty.Margin = Thickness(6)
            empty.Foreground = _brush("#A3A3A3")
            self.historyPanel.Children.Add(empty)
        else:
            auto_expand_rooms_left = 2
            for room_label in sorted(grouped_all.keys()):
                room_notes_all = grouped_all[room_label]
                room_notes_open = [n for n in room_notes_all if not bool(n.get("completed", False))]
                open_count = len(room_notes_open)

                resolved_badge = ""
                rid = room_display_to_id.get(room_label, "")
                if rid and bool(self._room_resolved_cache.get(rid, False)):
                    resolved_badge = " [CHECK]"

                exp = Expander()
                exp.Header = "%s%s (Open: %s, Total: %s)" % (room_label, resolved_badge, open_count, len(room_notes_all))
                exp.Margin = Thickness(0, 0, 0, 8)
                exp.Foreground = _brush("#F3F4F6")
                exp.Tag = {
                    "openNotes": room_notes_open,
                    "loaded": False
                }

                placeholder = TextBlock()
                placeholder.Text = "Expand to view notes"
                placeholder.Margin = Thickness(6, 2, 6, 6)
                placeholder.Foreground = _brush("#737373")
                exp.Content = placeholder

                exp.Expanded += self.on_room_expander_expanded
                exp.Collapsed += self.on_room_expander_collapsed

                if open_count > 0 and auto_expand_rooms_left > 0:
                    exp.IsExpanded = True
                    auto_expand_rooms_left -= 1
                else:
                    exp.IsExpanded = False

                self.room_expanders.append(exp)
                self.historyPanel.Children.Add(exp)

        completed_notes = [n for n in all_filtered_active if bool(n.get("completed", False))]
        completed_expander = Expander()
        completed_expander.Header = "Completed (%s)" % len(completed_notes)
        completed_expander.Margin = Thickness(0, 6, 0, 8)
        completed_expander.Foreground = _brush("#D4D4D8")
        completed_expander.IsExpanded = False
        completed_expander.Visibility = Visibility.Visible if self.show_completed else Visibility.Collapsed
        completed_stack = StackPanel()
        if self.show_completed:
            for note in completed_notes:
                completed_stack.Children.Add(self._make_note_card(note, deleted_mode=False))
        completed_expander.Content = completed_stack
        self.historyPanel.Children.Add(completed_expander)

        deleted_notes = self._filtered_notes(include_completed=True, include_deleted=True, pre_normalized=deleted_normalized)
        deleted_expander = Expander()
        deleted_expander.Header = "Deleted Items (%s, auto-purge 7 days)" % len(deleted_notes)
        deleted_expander.Margin = Thickness(0, 2, 0, 6)
        deleted_expander.Foreground = _brush("#D4D4D8")
        deleted_expander.IsExpanded = False
        deleted_expander.Visibility = Visibility.Visible if self.show_deleted else Visibility.Collapsed
        deleted_stack = StackPanel()
        if self.show_deleted:
            for note in deleted_notes:
                deleted_stack.Children.Add(self._make_note_card(note, deleted_mode=True))
        deleted_expander.Content = deleted_stack
        self.historyPanel.Children.Add(deleted_expander)

    def _save(self):
        ok = save_notes(self.store_path, self.doc.Title, self.all_notes, self.work_log_items, self.custom_assignees)
        if not ok:
            forms.alert("Could not write redline file. Check folder permissions for project folder or Documents/DatumNotes.", warn_icon=True)

    def _set_active_tab(self, tab_name):
        self.active_tab = tab_name if tab_name in CATEGORY_TAB_OPTIONS else "All"
        self._update_tab_visuals()
        try:
            self._render_history()
        except Exception as ex:
            forms.alert(
                "Could not refresh this tab.\n%s\n\n%s" % (_safe_text(ex), traceback.format_exc()),
                warn_icon=True
            )

    def _selected_assignee_value(self):
        selected = _safe_text(self.assignedToCombo.SelectedItem).strip()
        if selected == "Custom...":
            return _safe_text(self.assignedToCustomText.Text).strip()
        if selected.startswith("Custom: "):
            return selected.replace("Custom: ", "", 1).strip()
        if selected.startswith("---"):
            return ""
        return selected

    def _match_room_from_text(self, room_text):
        value = _safe_text(room_text).strip().lower()
        if value in ["unassigned", "unknown", "n/a", "na", "none", "-"]:
            return unassigned_room_bucket()
        if not value:
            return unassigned_room_bucket()

        def _norm(text):
            return re.sub(r"[^a-z0-9]+", " ", _safe_text(text).strip().lower()).strip()

        value_norm = _norm(value)

        # Pass 1: exact match against display/name/number (raw + normalized)
        for r in self.rooms:
            display = _safe_text(r.get("roomDisplay", ""))
            name = _safe_text(r.get("name", ""))
            number = _safe_text(r.get("number", ""))

            display_l = display.lower()
            name_l = name.lower()
            number_l = number.lower()

            if value == display_l or value == name_l or value == number_l:
                return r

            if value_norm and (value_norm == _norm(display) or value_norm == _norm(name) or value_norm == _norm(number)):
                return r

        # Pass 2: partial containment both directions (raw + normalized)
        partial_matches = []
        for r in self.rooms:
            display = _safe_text(r.get("roomDisplay", ""))
            name = _safe_text(r.get("name", ""))
            number = _safe_text(r.get("number", ""))

            display_l = display.lower()
            name_l = name.lower()
            number_l = number.lower()

            display_n = _norm(display)
            name_n = _norm(name)
            number_n = _norm(number)

            raw_hit = (
                (value in display_l or display_l in value) or
                (value in name_l or name_l in value) or
                (value in number_l or number_l in value)
            )

            norm_hit = False
            if value_norm:
                norm_hit = (
                    (value_norm in display_n or display_n in value_norm) or
                    (value_norm in name_n or name_n in value_norm) or
                    (value_norm in number_n or number_n in value_norm)
                )

            if raw_hit or norm_hit:
                # Prefer the shortest likely room label for cleaner fuzzy choice.
                score = len(display) if display else 9999
                partial_matches.append((score, r))

        if partial_matches:
            partial_matches.sort(key=lambda x: x[0])
            return partial_matches[0][1]

        return unassigned_room_bucket()

    def _open_views_with_room(self, room_elem):
        uidoc = revit.uidoc
        results = []
        seen = set()
        for ui_view in uidoc.GetOpenUIViews():
            try:
                view = self.doc.GetElement(ui_view.ViewId)
                if not view or getattr(view, "IsTemplate", False):
                    continue
                bbox = room_elem.get_BoundingBox(view)
                if not bbox:
                    continue
                label = "%s - %s" % (_safe_text(view.ViewType), _safe_text(view.Name))
                if label in seen:
                    continue
                seen.add(label)
                results.append((label, view, ui_view))
            except Exception:
                pass
        return results

    def _make_note_id(self):
        return "N%s" % datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")

    def _make_work_item_id(self):
        return "W%s" % datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")

    def _set_note_assignee(self, note_id, assignee_text):
        """Update assignee on raw note payload to avoid legacy key conflicts."""
        target_id = _safe_text(note_id)
        if not target_id:
            return False

        updated = False
        now_text = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        assignee_value = _safe_text(assignee_text).strip()

        for i, note in enumerate(self.all_notes):
            raw = note if isinstance(note, dict) else {}
            normalized = _normalize_note(raw)
            if _safe_text(normalized.get("id")) != target_id:
                continue

            raw["assignedTo"] = assignee_value
            raw["assigned_to"] = assignee_value
            if "assigned" in raw:
                try:
                    del raw["assigned"]
                except Exception:
                    pass
            if "assignee" in raw:
                try:
                    del raw["assignee"]
                except Exception:
                    pass

            raw["editedAt"] = now_text
            raw["editedBy"] = self.current_user
            self.all_notes[i] = raw
            updated = True

        return updated

    def _set_note_room(self, note_id, room_data):
        """Update room fields on raw note payload so tab/filter state updates correctly."""
        target_id = _safe_text(note_id)
        if not target_id or not isinstance(room_data, dict):
            return False

        updated = False
        now_text = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for i, note in enumerate(self.all_notes):
            raw = note if isinstance(note, dict) else {}
            normalized = _normalize_note(raw)
            if _safe_text(normalized.get("id")) != target_id:
                continue

            raw["roomId"] = _safe_text(room_data.get("roomId", ""))
            raw["roomDisplay"] = _safe_text(room_data.get("roomDisplay", ""))
            raw["roomNumber"] = _safe_text(room_data.get("number", ""))
            raw["roomName"] = _safe_text(room_data.get("name", ""))
            raw["level"] = _safe_text(room_data.get("level", ""))
            raw["elementId"] = _safe_text(room_data.get("elementId", ""))

            # Remove legacy aliases if present.
            if "room_id" in raw:
                try:
                    del raw["room_id"]
                except Exception:
                    pass
            if "room" in raw and isinstance(raw.get("room"), dict):
                try:
                    del raw["room"]
                except Exception:
                    pass

            raw["editedAt"] = now_text
            raw["editedBy"] = self.current_user
            self.all_notes[i] = raw
            updated = True

        return updated

    def _set_note_completed(self, note_id, checked):
        """Update completion state directly on raw note payload by id."""
        target_id = _safe_text(note_id)
        if not target_id:
            return False

        updated = False
        now_text = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        completed_at = now_text if checked else ""
        completed_by = self.current_user if checked else ""

        for i, note in enumerate(self.all_notes):
            raw = note if isinstance(note, dict) else {}
            normalized = _normalize_note(raw)
            if _safe_text(normalized.get("id")) != target_id:
                continue

            raw["completed"] = bool(checked)
            raw["completedAt"] = completed_at
            raw["completedBy"] = completed_by
            # Completing a note clears pending.
            if checked:
                raw["pending"] = False
            raw["editedAt"] = now_text
            raw["editedBy"] = self.current_user

            self.all_notes[i] = raw
            updated = True
            break  # IDs are unique; stop after finding the match.

        return updated

    def _set_note_pending(self, note_id, pending):
        """Update pending state on raw note payload by id."""
        target_id = _safe_text(note_id)
        if not target_id:
            return False

        updated = False
        now_text = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for i, note in enumerate(self.all_notes):
            raw = note if isinstance(note, dict) else {}
            normalized = _normalize_note(raw)
            if _safe_text(normalized.get("id")) != target_id:
                continue

            raw["pending"] = bool(pending)
            # Marking pending clears completed.
            if pending:
                raw["completed"] = False
                raw["completedAt"] = ""
                raw["completedBy"] = ""
            raw["editedAt"] = now_text
            raw["editedBy"] = self.current_user

            self.all_notes[i] = raw
            updated = True
            break

        return updated

    def _pick_bulk_note_ids(self, title_text):
        """Pick multiple notes from the current filtered active set."""
        notes = [_normalize_note(n) for n in self._filtered_notes(include_completed=True, include_deleted=False)]
        if not notes:
            forms.alert("No notes match the current filters.", warn_icon=True)
            return []

        labels = []
        lookup = {}
        for n in notes:
            note_id = _safe_text(n.get("id", ""))
            if not note_id:
                continue

            room_label = _safe_text(n.get("roomDisplay", "Unknown Room"))
            cat_label = _normalize_category(n.get("category", "Observation"))
            assigned = _safe_text(n.get("assignedTo", "")).strip() or "none"
            body = _safe_text(n.get("text", "")).replace("\n", " ").strip()
            if len(body) > 52:
                body = body[:49] + "..."
            label = "[%s] %s | %s | %s | %s" % (note_id[-6:], room_label, cat_label, assigned, body)
            labels.append(label)
            lookup[label] = note_id

        picked = forms.SelectFromList.show(labels, title=title_text, button_name="Apply", multiselect=True)
        if not picked:
            return []

        picked_ids = []
        for p in picked:
            pid = lookup.get(_safe_text(p), "")
            if pid:
                picked_ids.append(pid)
        return picked_ids

    def _describe_work_day(self, day_text):
        day_dt = _parse_note_datetime(day_text)
        if day_dt == datetime.datetime.min:
            return day_text or "Unknown Day"

        today = datetime.datetime.now().date()
        if day_dt.date() == today:
            return "Today"
        if day_dt.date() == (today - datetime.timedelta(days=1)):
            return "Yesterday"
        return day_dt.strftime("%b %d, %Y")

    def _work_items_sorted(self):
        items = []
        for raw_item in self.work_log_items:
            item = _normalize_work_item(raw_item)
            if item.get("text"):
                items.append(item)
        return sorted(items, key=lambda x: _parse_note_datetime(x.get("timestamp", "")), reverse=True)

    def _work_items_grouped_by_day(self, items):
        grouped = {}
        for item in items:
            day_text = _safe_text(item.get("day", "")).strip() or _safe_text(item.get("timestamp", ""))[:10]
            grouped.setdefault(day_text, []).append(item)
        return grouped

    def _selected_work_items(self):
        items = self._work_items_sorted()
        if self.selected_work_day:
            items = [x for x in items if _safe_text(x.get("day", "")) == self.selected_work_day]
        return items

    def _build_work_calendar_tooltip(self):
        items = self._work_items_sorted()
        if not items:
            return "No work history yet."

        grouped = self._work_items_grouped_by_day(items)
        lines = ["Hover summary of saved work days:"]
        for day_text in sorted(grouped.keys(), reverse=True)[:10]:
            day_items = grouped[day_text]
            lines.append("")
            lines.append("%s (%s)" % (self._describe_work_day(day_text), len(day_items)))
            for item in day_items[:3]:
                stamp = _safe_text(item.get("timestamp", ""))[11:16]
                body = _safe_text(item.get("text", "")).replace("\n", " ").strip()
                if len(body) > 72:
                    body = body[:69] + "..."
                lines.append("  %s  %s" % (stamp or "--:--", body))
            if len(day_items) > 3:
                lines.append("  +%s more" % (len(day_items) - 3))
        return "\n".join(lines)

    def _make_work_item_card(self, item):
        card = Border()
        card.Margin = Thickness(2, 0, 2, 8)
        card.Padding = Thickness(10)
        card.Background = _brush(THEME_CARD_BG)
        card.BorderBrush = _brush(THEME_PANEL_BORDER)
        card.BorderThickness = Thickness(1)

        content = StackPanel()

        top_grid = Grid()
        top_grid.ColumnDefinitions.Add(ColumnDefinition())
        top_grid.ColumnDefinitions.Add(ColumnDefinition())
        top_grid.ColumnDefinitions[1].Width = GridLength.Auto

        stamp = TextBlock()
        stamp.Text = _safe_text(item.get("timestamp", ""))
        stamp.Foreground = _brush("#D4D4D8")
        stamp.FontSize = 11
        Grid.SetColumn(stamp, 0)
        top_grid.Children.Add(stamp)

        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal

        edit_btn = Button()
        edit_btn.Content = "Edit"
        edit_btn.MinWidth = 52
        edit_btn.Height = 24
        edit_btn.Padding = Thickness(8, 1, 8, 1)
        edit_btn.Margin = Thickness(0, 0, 4, 0)
        edit_btn.Tag = item.get("id")
        edit_btn.Background = _brush(THEME_BUTTON_BG)
        edit_btn.Foreground = _brush("#F3F4F6")
        edit_btn.BorderBrush = _brush(THEME_BUTTON_BORDER)
        edit_btn.Click += self.on_edit_work_item
        button_panel.Children.Add(edit_btn)

        delete_btn = Button()
        delete_btn.Content = "Delete"
        delete_btn.MinWidth = 52
        delete_btn.Height = 24
        delete_btn.Padding = Thickness(8, 1, 8, 1)
        delete_btn.Tag = item.get("id")
        delete_btn.Background = _brush("#3F1D1D")
        delete_btn.Foreground = _brush("#FEE2E2")
        delete_btn.BorderBrush = _brush("#7F1D1D")
        delete_btn.Click += self.on_delete_work_item
        button_panel.Children.Add(delete_btn)

        Grid.SetColumn(button_panel, 1)
        top_grid.Children.Add(button_panel)

        content.Children.Add(top_grid)

        body = TextBlock()
        body.Text = _safe_text(item.get("text", ""))
        body.TextWrapping = TextWrapping.Wrap
        body.Margin = Thickness(0, 8, 0, 0)
        body.Foreground = _brush("#F3F4F6")
        content.Children.Add(body)

        edited_at = _safe_text(item.get("editedAt", "")).strip()
        if edited_at:
            meta = TextBlock()
            meta.Text = "Edited %s" % edited_at
            if _safe_text(item.get("editedBy", "")).strip():
                meta.Text = "%s by %s" % (meta.Text, _safe_text(item.get("editedBy", "")))
            meta.Margin = Thickness(0, 6, 0, 0)
            meta.Foreground = _brush("#C4C4C4")
            meta.FontSize = 11
            content.Children.Add(meta)

        card.Child = content
        return card

    def _add_work_day_group(self, host, day_text, items, expanded):
        expander = Expander()
        expander.Header = "%s (%s)" % (self._describe_work_day(day_text), len(items))
        expander.Margin = Thickness(0, 0, 0, 8)
        expander.Foreground = _brush("#F3F4F6")
        expander.IsExpanded = expanded

        stack = StackPanel()
        for item in items:
            stack.Children.Add(self._make_work_item_card(item))
        expander.Content = stack
        host.Children.Add(expander)

    def _render_work_history(self):
        self.workHistoryPanel.Children.Clear()
        self.calendarWorkButton.ToolTip = self._build_work_calendar_tooltip()

        items = self._selected_work_items()
        if self.selected_work_day:
            self.workHistoryMetaText.Text = "Viewing %s. Use Calendar to switch back to all days." % self._describe_work_day(self.selected_work_day)
        else:
            self.workHistoryMetaText.Text = "Recent work shows here. Items older than one week are grouped by default."

        if not items:
            empty = TextBlock()
            empty.Text = "No work history yet. Add an item above to start tracking today’s focus."
            empty.Margin = Thickness(6)
            empty.Foreground = _brush("#C4C4C4")
            self.workHistoryPanel.Children.Add(empty)
            return

        if self.selected_work_day:
            self._add_work_day_group(self.workHistoryPanel, self.selected_work_day, items, True)
            return

        grouped = self._work_items_grouped_by_day(items)
        cutoff = datetime.datetime.now() - datetime.timedelta(days=7)
        older_groups = []

        for day_text in sorted(grouped.keys(), reverse=True):
            parsed = _parse_note_datetime(day_text)
            is_recent = parsed != datetime.datetime.min and parsed >= cutoff
            if is_recent:
                self._add_work_day_group(self.workHistoryPanel, day_text, grouped[day_text], day_text == datetime.datetime.now().strftime("%Y-%m-%d"))
            else:
                older_groups.append((day_text, grouped[day_text]))

        if older_groups:
            older_expander = Expander()
            older_expander.Header = "Older Than One Week (%s days)" % len(older_groups)
            older_expander.Margin = Thickness(0, 4, 0, 8)
            older_expander.Foreground = _brush("#F3F4F6")
            older_expander.IsExpanded = False

            older_stack = StackPanel()
            for day_text, day_items in older_groups:
                self._add_work_day_group(older_stack, day_text, day_items, False)
            older_expander.Content = older_stack
            self.workHistoryPanel.Children.Add(older_expander)

    def on_category_changed(self, sender, args):
        self._toggle_due_date_visibility()

    def on_assigned_to_changed(self, sender, args):
        self._toggle_custom_assignee_ui()

    def on_add_custom_assignee(self, sender, args):
        value = _safe_text(self.assignedToCustomText.Text).strip()
        if not value:
            forms.alert("Enter a custom name first.", warn_icon=True)
            return

        self._remember_custom_assignee(value)
        self._save()

        self._bind_assignees()
        picked = "Custom: %s" % value
        if picked in [ _safe_text(x) for x in self.assignedToCombo.Items ]:
            self.assignedToCombo.SelectedItem = picked
        else:
            self.assignedToCombo.SelectedItem = "Custom..."
            self.assignedToCustomText.Text = value

    def on_filter_changed(self, sender, args):
        try:
            self._render_history()
        except Exception as ex:
            forms.alert("Could not refresh filter.\n%s\n\n%s" % (_safe_text(ex), traceback.format_exc()), warn_icon=True)

    def on_sort_changed(self, sender, args):
        try:
            self._render_history()
        except Exception as ex:
            forms.alert("Could not apply sort.\n%s\n\n%s" % (_safe_text(ex), traceback.format_exc()), warn_icon=True)

    def on_search_changed(self, sender, args):
        try:
            self._render_history()
        except Exception as ex:
            forms.alert("Could not refresh search.\n%s\n\n%s" % (_safe_text(ex), traceback.format_exc()), warn_icon=True)

    def on_tab_all(self, sender, args):
        self._set_active_tab("All")

    def on_tab_action(self, sender, args):
        self._set_active_tab("Action Item")

    def on_tab_question(self, sender, args):
        self._set_active_tab("Question")

    def on_tab_decision(self, sender, args):
        self._set_active_tab("Decision")

    def on_tab_observation(self, sender, args):
        self._set_active_tab("Observation")

    def on_tab_unassigned(self, sender, args):
        self._set_active_tab("Pending")

    def on_toggle_completed(self, sender, args):
        self.show_completed = not self.show_completed
        self._render_history()

    def on_toggle_deleted(self, sender, args):
        self.show_deleted = not self.show_deleted
        self._render_history()

    def on_expand_all(self, sender, args):
        for exp in self.room_expanders:
            exp.IsExpanded = True

    def on_collapse_all(self, sender, args):
        for exp in self.room_expanders:
            exp.IsExpanded = False

    def on_room_changed(self, sender, args):
        selected_label = _safe_text(self.roomCombo.SelectedItem)
        self.selected_room = self.room_lookup.get(selected_label)
        self._update_selected_room_ui()

    def on_toggle_room_resolved(self, sender, args):
        if not self.selected_room:
            forms.alert("Select a room first.", warn_icon=True)
            return

        room_id = _safe_text(self.selected_room.get("roomId"))
        if not room_id:
            return

        if room_id in self.manual_resolved_rooms:
            self.manual_resolved_rooms.remove(room_id)
        else:
            self.manual_resolved_rooms.add(room_id)

        self._save_config()
        self._bind_rooms()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()

    def on_go_to_room(self, sender, args):
        room = self.selected_room
        if not room:
            forms.alert("Select a room first.", warn_icon=True)
            return

        element_id_text = _safe_text(room.get("elementId"))
        if not element_id_text:
            forms.alert("Could not resolve room element ID.", warn_icon=True)
            return

        try:
            room_elem = self.doc.GetElement(DB.ElementId(int(element_id_text)))
            if not room_elem:
                forms.alert("Room element not found in current model.", warn_icon=True)
                return

            candidates = self._open_views_with_room(room_elem)
            if not candidates:
                forms.alert("No currently open views contain this room. Open a view with this room and try again.", warn_icon=True)
                return

            labels = [c[0] for c in candidates]
            selected = forms.SelectFromList.show(labels, title="Go To Room - Pick Open View", button_name="Go", multiselect=False)
            if not selected:
                return

            picked = None
            for c in candidates:
                if c[0] == selected:
                    picked = c
                    break
            if not picked:
                return

            uidoc = revit.uidoc
            view = picked[1]
            ui_view = picked[2]

            try:
                uidoc.RequestViewChange(view)
            except Exception:
                pass

            try:
                ui_view.ZoomToFit()
            except Exception:
                pass

            uidoc.ShowElements(room_elem.Id)
            ids = List[DB.ElementId]()
            ids.Add(room_elem.Id)
            uidoc.Selection.SetElementIds(ids)
        except Exception as ex:
            forms.alert("Could not navigate to room.\n%s" % _safe_text(ex), warn_icon=True)

    def on_work_text_keydown(self, sender, args):
        if args.Key == Key.Enter:
            args.Handled = True
            self.on_add_work_item(None, None)

    def on_note_text_keydown(self, sender, args):
        if args.Key == Key.Enter:
            args.Handled = True
            self.on_add_note(None, None)

    def on_window_keydown(self, sender, args):
        if Keyboard.Modifiers != ModifierKeys.Control:
            return

        if args.Key == Key.G:
            args.Handled = True
            self.on_go_to_room(None, None)
        elif args.Key == Key.E:
            args.Handled = True
            self.on_export(None, None)
        elif args.Key == Key.F:
            args.Handled = True
            self.historySearchBox.Focus()

    def on_add_work_item(self, sender, args):
        work_text = _safe_text(self.workTodayText.Text).strip()
        if not work_text:
            forms.alert("Enter what you want to work on before logging it.", warn_icon=True)
            return

        now = datetime.datetime.now()
        item = {
            "id": self._make_work_item_id(),
            "timestamp": now.strftime("%Y-%m-%d %H:%M:%S"),
            "day": now.strftime("%Y-%m-%d"),
            "text": work_text,
            "author": self.current_user,
            "editedAt": "",
            "editedBy": ""
        }

        self.work_log_items.append(item)
        self.selected_work_day = now.strftime("%Y-%m-%d")
        self.workTodayText.Text = ""
        self._save()
        self._render_work_history()

    def on_pick_work_day(self, sender, args):
        items = self._work_items_sorted()
        if not items:
            forms.alert("No saved work days yet.", warn_icon=True)
            return

        grouped = self._work_items_grouped_by_day(items)
        labels = ["All Days"]
        lookup = {}
        for day_text in sorted(grouped.keys(), reverse=True):
            label = "%s (%s)" % (self._describe_work_day(day_text), len(grouped[day_text]))
            lookup[label] = day_text
            labels.append(label)

        selected = forms.SelectFromList.show(labels, title="Work History Days", button_name="Show", multiselect=False)
        if not selected:
            return

        if selected == "All Days":
            self.selected_work_day = ""
        else:
            self.selected_work_day = lookup.get(_safe_text(selected), "")
        self._render_work_history()

    def on_edit_work_item(self, sender, args):
        item_id = _safe_text(sender.Tag)
        if not item_id:
            return

        for index, raw in enumerate(self.work_log_items):
            item = _normalize_work_item(raw)
            if _safe_text(item.get("id")) != item_id:
                continue

            updated_text = forms.ask_for_string(
                default=_safe_text(item.get("text", "")),
                prompt="Edit work history entry",
                title="Edit Work Item"
            )
            if updated_text is None:
                return

            updated_text = _safe_text(updated_text).strip()
            if not updated_text:
                forms.alert("Work history text cannot be empty.", warn_icon=True)
                return

            item["text"] = updated_text
            item["editedAt"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            item["editedBy"] = self.current_user
            self.work_log_items[index] = item
            self._save()
            self._render_work_history()
            return

    def on_delete_work_item(self, sender, args):
        item_id = _safe_text(sender.Tag)
        if not item_id:
            return

        confirmed = forms.alert("Are you sure you want to delete this work item?", yes=True, no=True)
        if not confirmed:
            return

        self.work_log_items = [item for item in self.work_log_items if _safe_text(_normalize_work_item(item).get("id")) != item_id]
        self._save()
        self._render_work_history()

    def on_add_note(self, sender, args):
        selected_label = _safe_text(self.roomCombo.SelectedItem)
        note_text = _safe_text(self.noteText.Text).strip()
        category = _normalize_category(self.categoryCombo.SelectedItem)
        assigned_to = self._selected_assignee_value()
        if _safe_text(self.assignedToCombo.SelectedItem).strip() == "Custom...":
            self._remember_custom_assignee(assigned_to)
            self._bind_assignees()

        due_date = ""
        if category == "Action Item":
            selected_due = self.dueDatePicker.SelectedDate
            if selected_due:
                due_date = "%04d-%02d-%02d" % (selected_due.Year, selected_due.Month, selected_due.Day)

        if not selected_label or selected_label not in self.room_lookup:
            forms.alert("Please select a room or General.", warn_icon=True)
            return

        if not note_text:
            forms.alert("Please enter a note before clicking Add Redline.", warn_icon=True)
            return

        now = datetime.datetime.now()
        room = self.room_lookup[selected_label]

        note_item = {
            "id": self._make_note_id(),
            "timestamp": now.strftime("%Y-%m-%d %H:%M:%S"),
            "roomId": room["roomId"],
            "roomDisplay": room["roomDisplay"],
            "roomNumber": room["number"],
            "roomName": room["name"],
            "level": room["level"],
            "elementId": room["elementId"],
            "text": note_text,
            "completed": False,
            "completedAt": "",
            "completedBy": "",
            "editedAt": "",
            "editedBy": "",
            "category": category,
            "assignedTo": assigned_to,
            "dueDate": due_date,
            "deleted": False,
            "deletedAt": "",
            "imported": False,
            "duplicateFrom": "",
            "comments": []
        }

        self.all_notes.append(note_item)

        self._save()
        self._bind_history_filter()
        self._update_selected_room_ui()

        self.noteText.Text = ""
        if _safe_text(self.assignedToCombo.SelectedItem).strip() != "Custom...":
            self.assignedToCustomText.Text = ""
        self.dueDatePicker.SelectedDate = None
        self._render_history()

    def on_note_toggled(self, sender, args):
        note_id = _safe_text(sender.Tag)
        checked = bool(sender.IsChecked)
        if not note_id:
            forms.alert("This item is missing an ID and cannot be updated. Reopen the tool and try again.", warn_icon=True)
            return

        if not self._set_note_completed(note_id, checked):
            forms.alert("Could not update completion for this note.", warn_icon=True)
            return

        self._save()
        self._update_selected_room_ui()
        self._render_history()

    def on_note_pending_toggled(self, sender, args):
        note_id = _safe_text(sender.Tag)
        pending = bool(sender.IsChecked)
        if not note_id:
            forms.alert("This item is missing an ID and cannot be updated. Reopen the tool and try again.", warn_icon=True)
            return

        if not self._set_note_pending(note_id, pending):
            forms.alert("Could not update pending state for this note.", warn_icon=True)
            return

        self._save()
        self._update_selected_room_ui()
        self._render_history()

    def on_note_pending_toggled(self, sender, args):
        note_id = _safe_text(sender.Tag)
        pending = bool(sender.IsChecked)
        if not note_id:
            forms.alert("This item is missing an ID and cannot be updated. Reopen the tool and try again.", warn_icon=True)
            return

        if not self._set_note_pending(note_id, pending):
            forms.alert("Could not update pending state for this note.", warn_icon=True)
            return

        self._save()
        self._update_selected_room_ui()
        self._render_history()

    def on_add_comment(self, sender, args):
        comment_input = sender.Tag
        if not comment_input:
            return

        note_id = _safe_text(getattr(comment_input, "Tag", ""))
        comment_text = _safe_text(getattr(comment_input, "Text", "")).strip()

        if not note_id:
            return
        if not comment_text:
            forms.alert("Enter a comment before adding.", warn_icon=True)
            return

        now_text = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for i, note in enumerate(self.all_notes):
            n = _normalize_note(note)
            if _safe_text(n.get("id")) == note_id:
                comments = n.get("comments", []) if isinstance(n.get("comments", []), list) else []
                comments.append({
                    "timestamp": now_text,
                    "author": self.current_user,
                    "text": comment_text
                })
                n["comments"] = comments
                self.all_notes[i] = n
                break

        comment_input.Text = ""
        self._save()
        self._render_history()

    def on_edit_note(self, sender, args):
        note_id = _safe_text(sender.Tag)
        if not note_id:
            return

        idx = -1
        current = None
        for i, note in enumerate(self.all_notes):
            n = _normalize_note(note)
            if _safe_text(n.get("id")) == note_id:
                idx = i
                current = n
                break

        if idx < 0 or not current:
            return

        updated_text = forms.ask_for_string(
            default=_safe_text(current.get("text", "")),
            prompt="Edit note text",
            title="Edit Redline"
        )
        if updated_text is None:
            return
        updated_text = _safe_text(updated_text).strip()
        if not updated_text:
            forms.alert("Note text cannot be empty.", warn_icon=True)
            return

        category = forms.SelectFromList.show(
            CATEGORY_OPTIONS,
            title="Edit Category",
            button_name="Apply",
            multiselect=False
        )
        if category is None:
            return
        category = _normalize_category(category)

        assigned_default = _safe_text(current.get("assignedTo", "")).strip()
        assigned_to = forms.ask_for_string(
            default=assigned_default,
            prompt="Assigned to (leave blank for none)",
            title="Edit Assigned To"
        )
        if assigned_to is None:
            return

        due_date = ""
        if category == "Action Item":
            due_default = _safe_text(current.get("dueDate", "")).strip()
            due_date = forms.ask_for_string(
                default=due_default,
                prompt="Due date (YYYY-MM-DD, optional)",
                title="Edit Due Date"
            )
            if due_date is None:
                return

        current_room_display = _safe_text(current.get("roomDisplay", "")).strip()
        room_options = ["Keep Current Room"]
        room_lookup = {"Keep Current Room": None}

        unassigned_label = UNASSIGNED_ROOM_DISPLAY
        room_options.append(unassigned_label)
        room_lookup[unassigned_label] = unassigned_room_bucket()

        for r in self.rooms:
            label = _safe_text(r.get("roomDisplay", "")).strip()
            if not label:
                continue
            if label not in room_lookup:
                room_options.append(label)
                room_lookup[label] = r

        room_prompt = "Choose room (current: %s)" % (current_room_display or "GENERAL")
        selected_room_label = forms.SelectFromList.show(
            room_options,
            title=room_prompt,
            button_name="Apply",
            multiselect=False
        )
        if selected_room_label is None:
            return

        selected_room = room_lookup.get(_safe_text(selected_room_label))

        current["text"] = updated_text
        current["category"] = category
        current["assignedTo"] = _safe_text(assigned_to).strip()
        current["assigned_to"] = current["assignedTo"]
        current["dueDate"] = _safe_text(due_date).strip() if category == "Action Item" else ""

        if selected_room is not None:
            current["roomId"] = _safe_text(selected_room.get("roomId", ""))
            current["roomDisplay"] = _safe_text(selected_room.get("roomDisplay", ""))
            current["roomNumber"] = _safe_text(selected_room.get("number", ""))
            current["roomName"] = _safe_text(selected_room.get("name", ""))
            current["level"] = _safe_text(selected_room.get("level", ""))
            current["elementId"] = _safe_text(selected_room.get("elementId", ""))

            # Keep raw payload in sync so tab/filter logic updates immediately.
            self._set_note_room(note_id, selected_room)

        current["editedAt"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        current["editedBy"] = self.current_user

        # Ensure raw payload keys are also synced for imported/legacy notes.
        self._set_note_assignee(note_id, current.get("assignedTo", ""))

        self.all_notes[idx] = current
        self._save()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()

    def on_reassign_note_room(self, sender, args):
        note_id = _safe_text(sender.Tag)
        if not note_id:
            return

        idx = -1
        current = None
        for i, note in enumerate(self.all_notes):
            n = _normalize_note(note)
            if _safe_text(n.get("id")) == note_id:
                idx = i
                current = n
                break

        if idx < 0 or not current:
            return

        current_room_display = _safe_text(current.get("roomDisplay", "")).strip() or "GENERAL"
        room_options = []
        room_lookup = {}

        unassigned_label = UNASSIGNED_ROOM_DISPLAY
        room_options.append(unassigned_label)
        room_lookup[unassigned_label] = unassigned_room_bucket()

        for r in self.rooms:
            label = _safe_text(r.get("roomDisplay", "")).strip()
            if not label:
                continue
            if label not in room_lookup:
                room_options.append(label)
                room_lookup[label] = r

        selected_room_label = forms.SelectFromList.show(
            room_options,
            title="Assign Room (current: %s)" % current_room_display,
            button_name="Assign",
            multiselect=False
        )
        if not selected_room_label:
            return

        selected_room = room_lookup.get(_safe_text(selected_room_label))
        if not selected_room:
            selected_room = unassigned_room_bucket()

        current["roomId"] = _safe_text(selected_room.get("roomId", ""))
        current["roomDisplay"] = _safe_text(selected_room.get("roomDisplay", ""))
        current["roomNumber"] = _safe_text(selected_room.get("number", ""))
        current["roomName"] = _safe_text(selected_room.get("name", ""))
        current["level"] = _safe_text(selected_room.get("level", ""))
        current["elementId"] = _safe_text(selected_room.get("elementId", ""))
        current["editedAt"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        current["editedBy"] = self.current_user

        # Write to raw payload so General tab/filter state updates correctly.
        self._set_note_room(note_id, selected_room)

        self.all_notes[idx] = current
        self._save()
        self._bind_history_filter()

        # If reassigned out of GENERAL while viewing that tab, jump to the new room view.
        target_room_id = _safe_text(selected_room.get("roomId", ""))
        target_room_display = _safe_text(selected_room.get("roomDisplay", ""))
        if self.active_tab == "Pending" and target_room_id != UNASSIGNED_ROOM_ID:
            self.active_tab = "All"
            self._update_tab_visuals()
            if target_room_display:
                self.historyFilterCombo.SelectedItem = target_room_display

        self._update_selected_room_ui()
        self._render_history()

    def on_reassign_note_assignee(self, sender, args):
        note_id = _safe_text(sender.Tag)
        if not note_id:
            return

        idx = -1
        current = None
        for i, note in enumerate(self.all_notes):
            n = _normalize_note(note)
            if _safe_text(n.get("id")) == note_id:
                idx = i
                current = n
                break

        if idx < 0 or not current:
            return

        current_assigned = _safe_text(current.get("assignedTo", "")).strip()
        options = ["Keep Current", "Clear Assigned"]
        for name in DEFAULT_TEAM_MEMBERS:
            if name not in options:
                options.append(name)
        for name in self.custom_assignees:
            if name and name not in options:
                options.append(name)
        options.append("Type Custom...")

        selected = forms.SelectFromList.show(
            options,
            title="Assign Person (current: %s)" % (current_assigned or "none"),
            button_name="Apply",
            multiselect=False
        )
        if selected is None:
            return

        selected_text = _safe_text(selected).strip()
        if selected_text == "Keep Current":
            return
        elif selected_text == "Clear Assigned":
            new_assigned = ""
        elif selected_text == "Type Custom...":
            typed = forms.ask_for_string(
                default=current_assigned,
                prompt="Assigned to (leave blank for none)",
                title="Assign Person"
            )
            if typed is None:
                return
            new_assigned = _safe_text(typed).strip()
            self._remember_custom_assignee(new_assigned)
        else:
            new_assigned = selected_text

        current["assignedTo"] = new_assigned
        current["assigned_to"] = new_assigned

        # Update raw payload keys so imported notes do not keep stale assignee values.
        if not self._set_note_assignee(note_id, new_assigned):
            current["editedAt"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            current["editedBy"] = self.current_user
            self.all_notes[idx] = current

        if new_assigned and new_assigned not in DEFAULT_TEAM_MEMBERS:
            self._remember_custom_assignee(new_assigned)

        self._save()
        self._bind_assignees()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()

    def on_reassign_note_category(self, sender, args):
        note_id = _safe_text(sender.Tag)
        if not note_id:
            return

        idx = -1
        current = None
        for i, note in enumerate(self.all_notes):
            n = _normalize_note(note)
            if _safe_text(n.get("id")) == note_id:
                idx = i
                current = n
                break

        if idx < 0 or not current:
            return

        current_category = _normalize_category(current.get("category", "Observation"))
        options = []
        for c in CATEGORY_OPTIONS:
            if c == current_category:
                options.append("%s (current)" % c)
            else:
                options.append(c)

        selected = forms.SelectFromList.show(
            options,
            title="Change Type",
            button_name="Apply",
            multiselect=False
        )
        if not selected:
            return

        selected_text = _safe_text(selected).replace(" (current)", "", 1).strip()
        new_category = _normalize_category(selected_text)
        if new_category == current_category:
            return

        current["category"] = new_category
        if new_category != "Action Item":
            current["dueDate"] = ""

        # If this note is currently General, offer room assignment so it leaves that bucket.
        current_room_id = _safe_text(current.get("roomId", ""))
        if current_room_id == UNASSIGNED_ROOM_ID:
            room_options = ["Keep General"]
            room_lookup = {"Keep General": None}
            for r in self.rooms:
                label = _safe_text(r.get("roomDisplay", "")).strip()
                if not label:
                    continue
                if label not in room_lookup:
                    room_options.append(label)
                    room_lookup[label] = r

            picked_room = forms.SelectFromList.show(
                room_options,
                title="This note is General. Move it to a room?",
                button_name="Apply",
                multiselect=False
            )
            if picked_room and _safe_text(picked_room) != "Keep General":
                chosen_room = room_lookup.get(_safe_text(picked_room))
                if chosen_room:
                    current["roomId"] = _safe_text(chosen_room.get("roomId", ""))
                    current["roomDisplay"] = _safe_text(chosen_room.get("roomDisplay", ""))
                    current["roomNumber"] = _safe_text(chosen_room.get("number", ""))
                    current["roomName"] = _safe_text(chosen_room.get("name", ""))
                    current["level"] = _safe_text(chosen_room.get("level", ""))
                    current["elementId"] = _safe_text(chosen_room.get("elementId", ""))
                    self._set_note_room(note_id, chosen_room)

        current["editedAt"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        current["editedBy"] = self.current_user
        self.all_notes[idx] = current

        self._save()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()

    def on_bulk_assign(self, sender, args):
        note_ids = self._pick_bulk_note_ids("Bulk Assign - Select Notes")
        if not note_ids:
            return

        options = ["Clear Assigned"]
        for name in DEFAULT_TEAM_MEMBERS:
            if name not in options:
                options.append(name)
        for name in self.custom_assignees:
            if name and name not in options:
                options.append(name)
        options.append("Type Custom...")

        selected = forms.SelectFromList.show(
            options,
            title="Bulk Assign - Choose Person",
            button_name="Apply",
            multiselect=False
        )
        if selected is None:
            return

        selected_text = _safe_text(selected).strip()
        if selected_text == "Clear Assigned":
            new_assigned = ""
        elif selected_text == "Type Custom...":
            typed = forms.ask_for_string(default="", prompt="Assigned to (leave blank for none)", title="Bulk Assign")
            if typed is None:
                return
            new_assigned = _safe_text(typed).strip()
            self._remember_custom_assignee(new_assigned)
        else:
            new_assigned = selected_text

        changed = 0
        for nid in note_ids:
            if self._set_note_assignee(nid, new_assigned):
                changed += 1

        if new_assigned and new_assigned not in DEFAULT_TEAM_MEMBERS:
            self._remember_custom_assignee(new_assigned)

        self._save()
        self._bind_assignees()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()
        forms.alert("Bulk assign updated %s notes." % changed)

    def on_bulk_type(self, sender, args):
        note_ids = self._pick_bulk_note_ids("Bulk Type - Select Notes")
        if not note_ids:
            return

        selected = forms.SelectFromList.show(
            CATEGORY_OPTIONS,
            title="Bulk Type - Choose Type",
            button_name="Apply",
            multiselect=False
        )
        if not selected:
            return

        new_category = _normalize_category(selected)
        now_text = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        changed = 0
        changed_unassigned_ids = []

        id_set = set(note_ids)
        for i, note in enumerate(self.all_notes):
            raw = note if isinstance(note, dict) else {}
            nid = _safe_text(_normalize_note(raw).get("id", ""))
            if not nid or nid not in id_set:
                continue

            raw["category"] = new_category
            if new_category != "Action Item":
                raw["dueDate"] = ""
            if _safe_text(raw.get("roomId", "")) == UNASSIGNED_ROOM_ID:
                changed_unassigned_ids.append(nid)
            raw["editedAt"] = now_text
            raw["editedBy"] = self.current_user
            self.all_notes[i] = raw
            changed += 1

        # Offer one-shot room assignment for all changed notes still in General.
        if changed_unassigned_ids:
            room_options = ["Keep General"]
            room_lookup = {"Keep General": None}
            for r in self.rooms:
                label = _safe_text(r.get("roomDisplay", "")).strip()
                if not label:
                    continue
                if label not in room_lookup:
                    room_options.append(label)
                    room_lookup[label] = r

            picked_room = forms.SelectFromList.show(
                room_options,
                title="%s changed notes are still General. Move them to one room?" % len(changed_unassigned_ids),
                button_name="Apply",
                multiselect=False
            )
            if picked_room and _safe_text(picked_room) != "Keep General":
                chosen_room = room_lookup.get(_safe_text(picked_room))
                if chosen_room:
                    for nid in changed_unassigned_ids:
                        self._set_note_room(nid, chosen_room)

        self._save()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()
        forms.alert("Bulk type updated %s notes." % changed)

    def on_delete_note(self, sender, args):
        note_id = _safe_text(sender.Tag)
        if not note_id:
            return

        confirmed = forms.alert("Are you sure you want to delete this note?", yes=True, no=True)
        if not confirmed:
            return

        now_text = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for i, note in enumerate(self.all_notes):
            n = _normalize_note(note)
            if _safe_text(n.get("id")) == note_id:
                n["deleted"] = True
                n["deletedAt"] = now_text
                self.all_notes[i] = n
                break

        self._save()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()

    def on_restore_note(self, sender, args):
        note_id = _safe_text(sender.Tag)
        if not note_id:
            return

        for i, note in enumerate(self.all_notes):
            n = _normalize_note(note)
            if _safe_text(n.get("id")) == note_id:
                n["deleted"] = False
                n["deletedAt"] = ""
                self.all_notes[i] = n
                break

        self._save()
        self._bind_history_filter()
        self._render_history()

    def on_hard_delete_note(self, sender, args):
        note_id = _safe_text(sender.Tag)
        if not note_id:
            return

        confirmed = forms.alert("Are you sure you want to permanently delete this note?", yes=True, no=True, warn_icon=True)
        if not confirmed:
            return

        self.all_notes = [n for n in self.all_notes if _safe_text(_normalize_note(n).get("id")) != note_id]

        self._save()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()

    def on_duplicate_note(self, sender, args):
        note_id = _safe_text(sender.Tag)
        if not note_id:
            return

        source = None
        for n in self.all_notes:
            nn = _normalize_note(n)
            if _safe_text(nn.get("id")) == note_id:
                source = nn
                break
        if not source:
            return

        room_labels = [self._room_label(r) for r in self.rooms]
        selected = forms.SelectFromList.show(room_labels, title="Duplicate To Room", button_name="Duplicate", multiselect=False)
        if not selected:
            return

        target_room = self.room_lookup.get(_safe_text(selected))
        if not target_room:
            return

        now = datetime.datetime.now()
        new_note = _normalize_note(source)
        new_note["id"] = self._make_note_id()
        new_note["timestamp"] = now.strftime("%Y-%m-%d %H:%M:%S")
        new_note["roomId"] = target_room["roomId"]
        new_note["roomDisplay"] = target_room["roomDisplay"]
        new_note["roomNumber"] = target_room["number"]
        new_note["roomName"] = target_room["name"]
        new_note["level"] = target_room["level"]
        new_note["elementId"] = target_room["elementId"]
        new_note["completed"] = False
        new_note["completedAt"] = ""
        new_note["completedBy"] = ""
        new_note["editedAt"] = ""
        new_note["editedBy"] = ""
        new_note["deleted"] = False
        new_note["deletedAt"] = ""
        new_note["duplicateFrom"] = _safe_text(source.get("roomDisplay", "Unknown Room"))
        if new_note.get("text"):
            new_note["text"] = "%s\n[Duplicated from %s]" % (new_note.get("text"), _safe_text(source.get("roomDisplay", "Unknown Room")))

        self.all_notes.append(new_note)
        self._save()
        self._bind_history_filter()
        self._render_history()

    def on_export(self, sender, args):
        default_export_path = html_path_for_document(self.doc)
        default_dir = _ensure_datum_notes_folder()
        default_name = "redline_export.html"

        if default_export_path:
            default_dir = os.path.dirname(default_export_path) or default_dir
            default_name = os.path.basename(default_export_path) or default_name

        save_dialog = SaveFileDialog()
        save_dialog.Title = "Export Redline"
        save_dialog.Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*"
        save_dialog.DefaultExt = ".html"
        save_dialog.AddExtension = True
        save_dialog.InitialDirectory = default_dir
        save_dialog.FileName = default_name

        selected = save_dialog.ShowDialog()
        if not selected:
            return

        export_path = _safe_text(save_dialog.FileName)
        if not export_path:
            forms.alert("No export location selected.", warn_icon=True)
            return

        export_settings = prompt_export_settings(self.all_notes)
        if export_settings is None:
            return

        try:
            meta = get_project_metadata(self.doc)
            html = build_export_html(meta["projectName"], meta["architectName"], self.all_notes, export_settings, meta.get("projectAddress", ""))
            with codecs.open(export_path, "w", "utf-8") as handle:
                handle.write(html)
        except Exception as ex:
            forms.alert("Export failed.\n%s" % _safe_text(ex), warn_icon=True)
            return

        forms.alert("Export complete:\n%s" % export_path)

    def on_weekly_digest_export(self, sender, args):
        default_dir = _ensure_datum_notes_folder()
        default_name = "redline_weekly_digest.html"

        save_dialog = SaveFileDialog()
        save_dialog.Title = "Export Weekly Digest"
        save_dialog.Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*"
        save_dialog.DefaultExt = ".html"
        save_dialog.AddExtension = True
        save_dialog.InitialDirectory = default_dir
        save_dialog.FileName = default_name

        selected = save_dialog.ShowDialog()
        if not selected:
            return

        export_path = _safe_text(save_dialog.FileName)
        if not export_path:
            forms.alert("No export location selected.", warn_icon=True)
            return

        try:
            meta = get_project_metadata(self.doc)
            html = build_weekly_digest_html(meta.get("projectName", "Untitled Project"), self.all_notes)
            with codecs.open(export_path, "w", "utf-8") as handle:
                handle.write(html)
        except Exception as ex:
            forms.alert("Weekly digest export failed.\n%s" % _safe_text(ex), warn_icon=True)
            return

        forms.alert("Weekly digest exported:\n%s" % export_path)

    def _prompt_team_export_filters(self):
        active = [_normalize_note(n) for n in self._active_notes()]
        if not active:
            return None

        assigned = sorted(list(set([_safe_text(n.get("assignedTo", "")).strip() for n in active if _safe_text(n.get("assignedTo", "")).strip()])))
        categories = sorted(list(set([_normalize_category(n.get("category", "Observation")) for n in active])))
        rooms = sorted(list(set([_safe_text(n.get("roomDisplay", "Unknown Room")) for n in active])))

        assigned_opt = forms.SelectFromList.show(["All"] + assigned, title="Team Export - Assigned To", button_name="Next", multiselect=False)
        if assigned_opt is None:
            return None
        category_opt = forms.SelectFromList.show(["All"] + categories, title="Team Export - Category", button_name="Next", multiselect=False)
        if category_opt is None:
            return None
        room_opt = forms.SelectFromList.show(["All"] + rooms, title="Team Export - Room", button_name="Export", multiselect=False)
        if room_opt is None:
            return None

        return {
            "assigned": _safe_text(assigned_opt),
            "category": _safe_text(category_opt),
            "room": _safe_text(room_opt)
        }

    def on_export_for_team_member(self, sender, args):
        filters = self._prompt_team_export_filters()
        if filters is None:
            return

        export_notes = []
        for note in self._active_notes():
            n = _normalize_note(note)
            if filters["assigned"] != "All" and _safe_text(n.get("assignedTo", "")).strip() != filters["assigned"]:
                continue
            if filters["category"] != "All" and _normalize_category(n.get("category", "Observation")) != filters["category"]:
                continue
            if filters["room"] != "All" and _safe_text(n.get("roomDisplay", "Unknown Room")) != filters["room"]:
                continue
            export_notes.append(n)

        if not export_notes:
            forms.alert("No notes match the selected export filters.", warn_icon=True)
            return

        save_dialog = SaveFileDialog()
        save_dialog.Title = "Export for Team Member"
        save_dialog.Filter = "Datum Notes (*.datumnotes)|*.datumnotes|JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        save_dialog.DefaultExt = ".datumnotes"
        save_dialog.AddExtension = True
        save_dialog.InitialDirectory = _ensure_datum_notes_folder()
        save_dialog.FileName = "redline_team_export.datumnotes"

        selected = save_dialog.ShowDialog()
        if not selected:
            return

        export_path = _safe_text(save_dialog.FileName)
        if not export_path:
            return

        payload = {
            "tool": "Redline",
            "exportedAt": datetime.datetime.now().isoformat(),
            "filters": filters,
            "redlineItems": export_notes
        }

        try:
            with codecs.open(export_path, "w", "utf-8") as handle:
                json.dump(payload, handle, indent=2)
        except Exception as ex:
            forms.alert("Export for team member failed.\n%s" % _safe_text(ex), warn_icon=True)
            return

        forms.alert("Team export complete:\n%s" % export_path)

    def on_import_team_file(self, sender, args):
        open_dialog = OpenFileDialog()
        open_dialog.Title = "Import Team File"
        open_dialog.Filter = "Datum Notes (*.datumnotes)|*.datumnotes|JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        open_dialog.InitialDirectory = _ensure_datum_notes_folder()

        selected = open_dialog.ShowDialog()
        if not selected:
            return

        import_path = _safe_text(open_dialog.FileName)
        if not import_path or not os.path.exists(import_path):
            forms.alert("Import file not found.", warn_icon=True)
            return

        try:
            with codecs.open(import_path, "r", "utf-8") as handle:
                payload = json.load(handle)
        except Exception as ex:
            forms.alert("Could not read import file.\n%s" % _safe_text(ex), warn_icon=True)
            return

        if isinstance(payload, dict):
            imported_list = payload.get("redlineItems", payload.get("notes", []))
        elif isinstance(payload, list):
            imported_list = payload
        else:
            imported_list = []

        existing_ids = set([_safe_text(_normalize_note(n).get("id")) for n in self.all_notes])
        added = 0
        skipped = 0

        for raw in imported_list:
            n = _normalize_note(raw)
            nid = _safe_text(n.get("id"))
            if not nid or nid in existing_ids:
                skipped += 1
                continue
            n["imported"] = True
            self.all_notes.append(n)
            existing_ids.add(nid)
            added += 1

        self._save()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()

        forms.alert("Import complete. Added: %s | Skipped duplicates: %s" % (added, skipped))

    def on_copy_ai_template(self, sender, args):
        meta = get_project_metadata(self.doc)
        rooms = sorted(list(set([
            _safe_text(_normalize_note(n).get("roomName", "")) or _safe_text(_normalize_note(n).get("roomDisplay", ""))
            for n in self._active_notes()
            if _safe_text(_normalize_note(n).get("roomName", "")) or _safe_text(_normalize_note(n).get("roomDisplay", ""))
        ])))
        template = build_ai_template(meta.get("projectName", "Untitled Project"), rooms)

        try:
            Clipboard.SetText(template)
        except Exception as ex:
            forms.alert("Could not copy template to clipboard.\n%s" % _safe_text(ex), warn_icon=True)
            return

        forms.alert("Template copied - paste it into your AI assistant with your transcript.")

    def on_import_from_ai(self, sender, args):
        pasted = forms.ask_for_string(default="", prompt="Paste AI output in the exact Redline template format", title="Import from AI")
        if pasted is None:
            return

        parsed = parse_ai_template_input(pasted)
        if not parsed:
            forms.alert("No valid entries found. Ensure the format uses room | note | ...", warn_icon=True)
            return

        added = 0
        now = datetime.datetime.now()
        for item in parsed:
            if bool(item.get("isUnassigned", False)):
                room = unassigned_room_bucket()
            else:
                room = self._match_room_from_text(item.get("room", ""))

            # Never drop imported AI items because of room matching failures.
            if not room:
                room = unassigned_room_bucket()

            if not room:
                continue

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

        for item in parsed:
            assigned = _safe_text(item.get("assignedTo", "")).strip()
            if assigned and assigned not in DEFAULT_TEAM_MEMBERS:
                self._remember_custom_assignee(assigned)

        self._save()
        self._bind_history_filter()
        self._bind_assignees()
        self._update_selected_room_ui()
        self._render_history()

        forms.alert("Imported %s notes from AI template." % added)

    def on_open_tutorial(self, sender, args):
        try:
            Process.Start(TUTORIAL_URL)
        except Exception:
            forms.alert("Could not open browser. Visit: %s" % TUTORIAL_URL, warn_icon=True)

    def on_open_upload_info(self, sender, args):
        try:
            Process.Start(UPLOAD_INFO_URL)
        except Exception:
            forms.alert("Could not open browser. Visit: %s" % UPLOAD_INFO_URL, warn_icon=True)

    def on_open_update_info(self, sender, args):
        try:
            Process.Start(UPDATE_INFO_URL)
        except Exception:
            forms.alert("Could not open browser. Visit: %s" % UPDATE_INFO_URL, warn_icon=True)


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active Revit document.", warn_icon=True)
        return

    ensure_icon_png()

    xaml = script.get_bundle_file("redline_ui.xaml")
    if not xaml or not os.path.exists(xaml):
        forms.alert("UI file not found: redline_ui.xaml", warn_icon=True)
        return

    window = RedlineWindow(xaml, doc)
    window.ShowDialog()


if __name__ == "__main__":
    main()