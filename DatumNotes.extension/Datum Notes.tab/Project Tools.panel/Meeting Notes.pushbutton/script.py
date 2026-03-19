# -*- coding: utf-8 -*-
from __future__ import print_function

import os
import re
import json
import codecs
import base64
import datetime

from pyrevit import revit, DB, forms, script

from System.Windows import Thickness, Visibility, VerticalAlignment, TextWrapping, FontWeights, HorizontalAlignment, GridLength
from System.Windows.Controls import TextBlock, StackPanel, CheckBox, Border, Orientation, Button, Grid, ColumnDefinition, Expander, TextBox
from System.Diagnostics import Process
from Microsoft.Win32 import SaveFileDialog, OpenFileDialog
from System.Windows.Media import BrushConverter
from System.Collections.Generic import List
from System.Windows import Clipboard
from System.Windows.Input import Key, ModifierKeys, Keyboard


UPLOAD_INFO_URL = "https://datumnotes.com/from-revit"
CATEGORY_OPTIONS = ["Decision", "Action Item", "Question", "Observation"]
BADGE_COLORS = {
    "Decision": "#2563EB",
    "Action Item": "#DC2626",
    "Question": "#F59E0B",
    "Observation": "#6B7280"
}
DEFAULT_TEAM_MEMBERS = ["PM", "Architect", "Coordinator", "MEP", "Owner"]
SORT_OPTIONS = ["Newest First", "Oldest First", "By Room"]
CATEGORY_TAB_OPTIONS = ["All", "Action Item", "Question", "Decision", "Observation", "Unassigned"]
UNASSIGNED_ROOM_ID = "UNASSIGNED"
UNASSIGNED_ROOM_DISPLAY = "UNASSIGNED | Unassigned"
ICON_PNG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAYAAADimHc4AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAHjSURBVHhe7dExbsNQEANR9+5ygpw8t0ydNHJDwAB3IYZQNMVr7IX9iXk8Pz5/0PPQD/C3CFBGgDIClBGgjABlBCgjQBkByghQRoAyApQRoOy0AN9fz1vR/VsEWNL9WwRY0v1bBFjS/VuxAPr91aX2EcCU2kcAU2ofAUypfQQwpfYRwJTaRwBTah8BTKl9BDCl9hHAlNpXC6D3bfo+Nb13EeCg71PTexcBDvo+Nb13EeCg71PTe1ctwNWk9hHAlNpHAFNqHwFMqX0EMKX2EcCU2kcAU2ofAUypfbUAet+m71PTexcBDvo+Nb13EeCg71PTexcBDvo+Nb131QJcTWofAUypfQQwpfYRwJTaRwBTah8BTKl9BDCl9hHAlNpXC6D3afr/U2f/3gsBTGf/3svtAujn+v0703sXAeT7d6b3rlqAFn2n+97pvYsA5nun9y4CmO+d3rsIYL53eu+6XYCt1D4CmFL7CGBK7SOAKbWPAKbUPgKYUvsIYErtI4AptY8AptQ+AphS+whgSu0jgCm1jwCm1D4CmFL7CGBK7SOAKbWPAKbUPgKYUvtiAf473b9FgCXdv0WAJd2/RYAl3b91WgDsEKCMAGUEKCNAGQHKCFBGgDIClBGgjABlBCgjQBkByn4BQpL38pWaWnMAAAAASUVORK5CYII="


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
        "roomId": _safe_text(base.get("roomId")),
        "roomNumber": _safe_text(base.get("roomNumber")),
        "roomName": _safe_text(base.get("roomName")),
        "roomDisplay": _safe_text(base.get("roomDisplay") or "Unknown Room"),
        "level": _safe_text(base.get("level")),
        "elementId": _safe_text(base.get("elementId")),
        "text": _safe_text(base.get("text")),
        "completed": bool(base.get("completed", False)),
        "completedAt": _safe_text(base.get("completedAt")),
        "completedBy": _safe_text(base.get("completedBy")),
        "editedAt": _safe_text(base.get("editedAt")),
        "editedBy": _safe_text(base.get("editedBy")),
        "category": category,
        "assignedTo": _safe_text(base.get("assignedTo")),
        "dueDate": _safe_text(base.get("dueDate")) if category == "Action Item" else "",
        "deleted": bool(base.get("deleted", False)),
        "deletedAt": _safe_text(base.get("deletedAt")),
        "imported": bool(base.get("imported", False)),
        "duplicateFrom": _safe_text(base.get("duplicateFrom")),
        "comments": normalized_comments
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
        "manualResolvedRooms": []
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
        out["customAssignees"] = [x for x in custom_vals if _safe_text(x).strip()] if isinstance(custom_vals, list) else []
        out["manualResolvedRooms"] = [x for x in resolved_vals if _safe_text(x).strip()] if isinstance(resolved_vals, list) else []
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
        "level": "Unassigned",
        "number": "UNASSIGNED",
        "name": "Unassigned"
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


def load_notes(path):
    if not path or not os.path.exists(path):
        return []

    try:
        with codecs.open(path, "r", "utf-8") as handle:
            data = json.load(handle)
    except Exception:
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


def save_notes(path, doc_title, notes):
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

    payload = {
        "redlineProject": _safe_text(doc_title),
        "redlineSavedAt": datetime.datetime.now().isoformat(),
        "redlineItems": normalized_notes,
        "commentsByNoteId": comments_by_note
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
    return """Parse the transcript below and fill in this template exactly. For each item identify the room it relates to if possible. If a room cannot be determined place the item under a category called UNASSIGNED. Do not leave any action items questions or decisions out.
PROJECT: %s
DATE: %s
ACTION ITEMS: room | note | assigned to | due date
DECISIONS: room | note
QUESTIONS: room | note | assigned to
OBSERVATIONS: room | note
UNASSIGNED: note | category
""" % (project_name, date_text)


def parse_ai_template_input(text):
    lines = [_safe_text(x).rstrip() for x in _safe_text(text).splitlines()]
    parsed = []
    current = ""

    for line in lines:
        striped = line.strip()
        if not striped:
            continue

        upper = striped.upper()
        if upper.startswith("ACTION ITEMS:"):
            current = "ACTION ITEMS"
            inline = striped[len("ACTION ITEMS:"):].strip()
            if inline:
                striped = inline
            else:
                continue
        if upper.startswith("DECISIONS:"):
            current = "DECISIONS"
            inline = striped[len("DECISIONS:"):].strip()
            if inline:
                striped = inline
            else:
                continue
        if upper.startswith("QUESTIONS:"):
            current = "QUESTIONS"
            inline = striped[len("QUESTIONS:"):].strip()
            if inline:
                striped = inline
            else:
                continue
        if upper.startswith("OBSERVATIONS:"):
            current = "OBSERVATIONS"
            inline = striped[len("OBSERVATIONS:"):].strip()
            if inline:
                striped = inline
            else:
                continue
        if upper.startswith("UNASSIGNED:"):
            current = "UNASSIGNED"
            inline = striped[len("UNASSIGNED:"):].strip()
            if inline:
                striped = inline
            else:
                continue

        if current not in ["ACTION ITEMS", "DECISIONS", "QUESTIONS", "OBSERVATIONS", "UNASSIGNED"]:
            continue

        parts = [p.strip() for p in striped.split("|")]
        assigned = ""
        due = ""
        is_unassigned = False

        if current == "UNASSIGNED":
            if len(parts) < 1:
                continue
            room = "UNASSIGNED"
            note = parts[0]
            category_hint = parts[1] if len(parts) > 1 else "Observation"
            category = _normalize_category(category_hint)
            is_unassigned = True
        else:
            if len(parts) < 2:
                continue
            room = parts[0]
            note = parts[1]

            if current == "ACTION ITEMS":
                category = "Action Item"
                assigned = parts[2] if len(parts) > 2 else ""
                due = parts[3] if len(parts) > 3 else ""
            elif current == "DECISIONS":
                category = "Decision"
            elif current == "QUESTIONS":
                category = "Question"
                assigned = parts[2] if len(parts) > 2 else ""
            else:
                category = "Observation"

        if note:
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

        self.config = load_redline_config()
        self.custom_assignees = sorted(list(set([_safe_text(x).strip() for x in self.config.get("customAssignees", []) if _safe_text(x).strip()])))
        self.manual_resolved_rooms = set([_safe_text(x) for x in self.config.get("manualResolvedRooms", []) if _safe_text(x)])

        self.store_path = json_path_for_document(doc)
        self.all_notes = load_notes(self.store_path)

        if self._purge_deleted_notes():
            self._save()

        self._wire_events()
        self._bind_rooms()
        self._bind_categories()
        self._bind_assignees()
        self._bind_history_sort()
        self._bind_history_filter()
        self._update_tab_visuals()
        self._update_toggle_button_text()
        self._toggle_due_date_visibility()
        self._update_selected_room_ui()
        self._render_history()

    def _save_config(self):
        self.config["customAssignees"] = list(self.custom_assignees)
        self.config["manualResolvedRooms"] = sorted(list(self.manual_resolved_rooms))
        save_redline_config(self.config)

    def _wire_events(self):
        self.addNoteButton.Click += self.on_add_note
        self.exportButton.Click += self.on_export
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
        return [n for n in self.all_notes if not bool(_normalize_note(n).get("deleted", False))]

    def _deleted_notes(self):
        return [n for n in self.all_notes if bool(_normalize_note(n).get("deleted", False))]

    def _room_is_resolved(self, room):
        if not room:
            return False

        room_id = _safe_text(room.get("roomId"))
        if room_id in self.manual_resolved_rooms:
            return True

        room_notes = [
            _normalize_note(n) for n in self._active_notes()
            if _safe_text(_normalize_note(n).get("roomId")) == room_id
        ]
        action_items = [n for n in room_notes if _normalize_category(n.get("category")) == "Action Item"]
        if not action_items:
            return False

        for item in action_items:
            if not bool(item.get("completed", False)):
                return False
        return True

    def _room_label(self, room):
        base = _safe_text(room.get("roomDisplay", "Unknown Room"))
        if self._room_is_resolved(room):
            return "[CHECK] %s" % base
        return base

    def _bind_rooms(self):
        self.roomCombo.Items.Clear()
        self.room_lookup = {}

        if not self.rooms:
            self.roomCombo.Items.Add("No rooms found")
            self.roomCombo.SelectedIndex = 0
            self.roomCombo.IsEnabled = False
            return

        for room in self.rooms:
            label = self._room_label(room)
            self.room_lookup[label] = room
            self.roomCombo.Items.Add(label)

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
            self.assignedToCombo.Items.Add("--- Custom Names ---")
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

    def _filtered_notes(self, include_completed=True, include_deleted=False):
        selected_room = _safe_text(self.historyFilterCombo.SelectedItem)
        search_text = _safe_text(self.historySearchBox.Text).strip().lower()

        if include_deleted:
            notes = [_normalize_note(n) for n in self._deleted_notes()]
        else:
            notes = [_normalize_note(n) for n in self._active_notes()]

        if selected_room and selected_room != "All Rooms":
            notes = [n for n in notes if (_safe_text(n.get("roomDisplay", "Unknown Room")) or "Unknown Room") == selected_room]

        if self.active_tab != "All":
            if self.active_tab == "Unassigned":
                notes = [n for n in notes if _safe_text(n.get("roomId")) == UNASSIGNED_ROOM_ID]
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
        self._render_history()

    def _set_tab_button_style(self, button, is_active):
        button.Background = _brush("#F59E0B") if is_active else _brush("#262626")
        button.Foreground = _brush("#111111") if is_active else _brush("#F3F4F6")
        button.BorderBrush = _brush("#D97706") if is_active else _brush("#3F3F46")

    def _update_tab_visuals(self):
        self._set_tab_button_style(self.tabAllButton, self.active_tab == "All")
        self._set_tab_button_style(self.tabActionButton, self.active_tab == "Action Item")
        self._set_tab_button_style(self.tabQuestionButton, self.active_tab == "Question")
        self._set_tab_button_style(self.tabDecisionButton, self.active_tab == "Decision")
        self._set_tab_button_style(self.tabObservationButton, self.active_tab == "Observation")
        self._set_tab_button_style(self.tabUnassignedButton, self.active_tab == "Unassigned")

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

    def _make_note_card(self, note, deleted_mode=False):
        category = _normalize_category(note.get("category", "Observation"))
        comments_count = len(note.get("comments", [])) if isinstance(note.get("comments", []), list) else 0

        card = Border()
        card.Margin = Thickness(2, 0, 2, 8)
        card.Padding = Thickness(10)
        card.Background = _brush("#111111")
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

            edit_btn = Button()
            edit_btn.Content = "E"
            edit_btn.Width = 22
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
            dup_btn.Content = "D"
            dup_btn.Width = 22
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
            delete_btn.Content = "X"
            delete_btn.Width = 22
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

            comments_stack = StackPanel()

            add_row = Grid()
            add_row.ColumnDefinitions.Add(ColumnDefinition())
            add_row.ColumnDefinitions.Add(ColumnDefinition())
            add_row.ColumnDefinitions[1].Width = GridLength.Auto

            comment_input = TextBox()
            comment_input.MinHeight = 28
            comment_input.Margin = Thickness(0, 0, 8, 0)
            comment_input.Background = _brush("#0D0D0D")
            comment_input.Foreground = _brush("#F3F4F6")
            comment_input.BorderBrush = _brush("#3F3F46")
            comment_input.Tag = note.get("id")
            Grid.SetColumn(comment_input, 0)
            add_row.Children.Add(comment_input)

            add_comment_btn = Button()
            add_comment_btn.Content = "Add Comment"
            add_comment_btn.Tag = comment_input
            add_comment_btn.Padding = Thickness(8, 3, 8, 3)
            add_comment_btn.Background = _brush("#262626")
            add_comment_btn.Foreground = _brush("#F3F4F6")
            add_comment_btn.BorderBrush = _brush("#3F3F46")
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
                item_border.Background = _brush("#171717")
                item_border.BorderBrush = _brush("#262626")
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

            comments_expander.Content = comments_stack
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
            box = CheckBox()
            box.Margin = Thickness(0, 2, 0, 0)
            box.Content = "Mark complete"
            box.IsChecked = bool(note.get("completed", False))
            box.Tag = note.get("id")
            box.Foreground = _brush("#D4D4D8")
            box.Checked += self.on_note_toggled
            box.Unchecked += self.on_note_toggled
            content.Children.Add(box)

        card.Child = content
        return card

    def _render_history(self):
        self.historyPanel.Children.Clear()
        self.room_expanders = []
        if self._purge_deleted_notes():
            self._save()
            self._bind_history_filter()

        total, action_items, completed = self._global_stats()
        self.historyGlobalSummaryText.Text = "Total notes: %s | Action items: %s | Completed: %s" % (total, action_items, completed)
        self._update_toggle_button_text()

        all_filtered_active = self._filtered_notes(include_completed=True, include_deleted=False)
        grouped_all = self._group_notes_by_room(all_filtered_active)

        if not grouped_all:
            empty = TextBlock()
            empty.Text = "No redlines match the current filters."
            empty.Margin = Thickness(6)
            empty.Foreground = _brush("#A3A3A3")
            self.historyPanel.Children.Add(empty)
        else:
            for room_label in sorted(grouped_all.keys()):
                room_notes_all = grouped_all[room_label]
                room_notes_open = [n for n in room_notes_all if not bool(n.get("completed", False))]
                open_count = len(room_notes_open)

                resolved_badge = ""
                room_obj = None
                for r in self.rooms:
                    if _safe_text(r.get("roomDisplay")) == room_label:
                        room_obj = r
                        break
                if room_obj and self._room_is_resolved(room_obj):
                    resolved_badge = " [CHECK]"

                exp = Expander()
                exp.Header = "%s%s (Open: %s, Total: %s)" % (room_label, resolved_badge, open_count, len(room_notes_all))
                exp.Margin = Thickness(0, 0, 0, 8)
                exp.Foreground = _brush("#F3F4F6")
                exp.IsExpanded = True if open_count > 0 else False

                room_stack = StackPanel()
                if room_notes_open:
                    for note in room_notes_open:
                        room_stack.Children.Add(self._make_note_card(note, deleted_mode=False))
                else:
                    empty_room = TextBlock()
                    empty_room.Text = "No open items in this room."
                    empty_room.Foreground = _brush("#737373")
                    empty_room.Margin = Thickness(6, 2, 6, 6)
                    room_stack.Children.Add(empty_room)

                exp.Content = room_stack
                self.room_expanders.append(exp)
                self.historyPanel.Children.Add(exp)

        completed_notes = [n for n in self._filtered_notes(include_completed=True, include_deleted=False) if bool(n.get("completed", False))]
        completed_expander = Expander()
        completed_expander.Header = "Completed (%s)" % len(completed_notes)
        completed_expander.Margin = Thickness(0, 6, 0, 8)
        completed_expander.Foreground = _brush("#D4D4D8")
        completed_expander.IsExpanded = False
        completed_expander.Visibility = Visibility.Visible if self.show_completed else Visibility.Collapsed
        completed_stack = StackPanel()
        for note in completed_notes:
            completed_stack.Children.Add(self._make_note_card(note, deleted_mode=False))
        completed_expander.Content = completed_stack
        self.historyPanel.Children.Add(completed_expander)

        deleted_notes = self._filtered_notes(include_completed=True, include_deleted=True)
        deleted_expander = Expander()
        deleted_expander.Header = "Deleted Items (%s, auto-purge 7 days)" % len(deleted_notes)
        deleted_expander.Margin = Thickness(0, 2, 0, 6)
        deleted_expander.Foreground = _brush("#D4D4D8")
        deleted_expander.IsExpanded = False
        deleted_expander.Visibility = Visibility.Visible if self.show_deleted else Visibility.Collapsed
        deleted_stack = StackPanel()
        for note in deleted_notes:
            deleted_stack.Children.Add(self._make_note_card(note, deleted_mode=True))
        deleted_expander.Content = deleted_stack
        self.historyPanel.Children.Add(deleted_expander)

    def _save(self):
        ok = save_notes(self.store_path, self.doc.Title, self.all_notes)
        if not ok:
            forms.alert("Could not write redline file. Check folder permissions for project folder or Documents/DatumNotes.", warn_icon=True)

    def _set_active_tab(self, tab_name):
        self.active_tab = tab_name if tab_name in CATEGORY_TAB_OPTIONS else "All"
        self._update_tab_visuals()
        self._render_history()

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

        for r in self.rooms:
            display = _safe_text(r.get("roomDisplay", "")).lower()
            name = _safe_text(r.get("name", "")).lower()
            number = _safe_text(r.get("number", "")).lower()
            if value == display or value == name or value == number:
                return r
            if value in display:
                return r

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

    def on_category_changed(self, sender, args):
        self._toggle_due_date_visibility()

    def on_assigned_to_changed(self, sender, args):
        self._toggle_custom_assignee_ui()

    def on_add_custom_assignee(self, sender, args):
        value = _safe_text(self.assignedToCustomText.Text).strip()
        if not value:
            forms.alert("Enter a custom name first.", warn_icon=True)
            return

        if value not in self.custom_assignees:
            self.custom_assignees.append(value)
            self.custom_assignees = sorted(list(set(self.custom_assignees)))
            self._save_config()

        self._bind_assignees()
        self.assignedToCombo.SelectedItem = "Custom: %s" % value

    def on_filter_changed(self, sender, args):
        self._render_history()

    def on_sort_changed(self, sender, args):
        self._render_history()

    def on_search_changed(self, sender, args):
        self._render_history()

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
        self._set_active_tab("Unassigned")

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

    def on_add_note(self, sender, args):
        if not self.rooms:
            forms.alert("No rooms found in this model.", warn_icon=True)
            return

        selected_label = _safe_text(self.roomCombo.SelectedItem)
        note_text = _safe_text(self.noteText.Text).strip()
        category = _normalize_category(self.categoryCombo.SelectedItem)
        assigned_to = self._selected_assignee_value()

        due_date = ""
        if category == "Action Item":
            selected_due = self.dueDatePicker.SelectedDate
            if selected_due:
                due_date = "%04d-%02d-%02d" % (selected_due.Year, selected_due.Month, selected_due.Day)

        if not selected_label or selected_label not in self.room_lookup:
            forms.alert("Please select a room.", warn_icon=True)
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
        self.assignedToCustomText.Text = ""
        self.dueDatePicker.SelectedDate = None
        self._render_history()

    def on_note_toggled(self, sender, args):
        note_id = _safe_text(sender.Tag)
        checked = bool(sender.IsChecked)
        completed_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") if checked else ""
        completed_by = self.current_user if checked else ""

        for i, note in enumerate(self.all_notes):
            n = _normalize_note(note)
            if _safe_text(n.get("id")) == note_id:
                n["completed"] = checked
                n["completedAt"] = completed_at
                n["completedBy"] = completed_by
                self.all_notes[i] = n
                break

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

        current["text"] = updated_text
        current["category"] = category
        current["assignedTo"] = _safe_text(assigned_to).strip()
        current["dueDate"] = _safe_text(due_date).strip() if category == "Action Item" else ""
        current["editedAt"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        current["editedBy"] = self.current_user

        self.all_notes[idx] = current
        self._save()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()

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

        self._save()
        self._bind_history_filter()
        self._update_selected_room_ui()
        self._render_history()

        forms.alert("Imported %s notes from AI template." % added)

    def on_open_upload_info(self, sender, args):
        try:
            Process.Start(UPLOAD_INFO_URL)
        except Exception:
            forms.alert("Could not open browser. Visit: %s" % UPLOAD_INFO_URL, warn_icon=True)


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