# -*- coding: utf-8 -*-
import Autodesk
from pyrevit import revit, DB, forms, script
from pyrevit.compat import get_elementid_value_func, get_elementid_from_value_func
from pyrevit.interop import xl
import clr
import os
import System
from Autodesk.Revit.UI import TaskDialog
# Add Windows Forms references
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Form, Label, ListBox, Button, DialogResult, TextBox,
    FormBorderStyle, FormStartPosition, SelectionMode
)
from System import Array
from System.Drawing import Point, Size

# Global variable to store the last exported directory
last_export_dir = os.path.expanduser("~") + "\\Desktop"

# Revit compatibility helpers
get_eid_value = get_elementid_value_func()
get_elementid_from_value = get_elementid_from_value_func()


# -----------------------------------------------------------------------------
# Simple dialog for opening exported file
# -----------------------------------------------------------------------------

class OpenFileDialog(Form):
    def __init__(self, filepath):
        self.filepath = filepath
        self.Text = "Export Successful"
        self.Width = 300
        self.Height = 150
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        self.label = Label()
        self.label.Text = "Schedule exported successfully.\nWould you like to open the file?"
        self.label.Location = Point(20, 10)
        self.label.Size = Size(self.ClientSize.Width - 40, 40)
        self.label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter

        self.open_button = Button()
        self.open_button.Text = "Open File"
        button_width = 100
        self.open_button.Location = Point((self.ClientSize.Width - button_width) / 2, 80)
        self.open_button.Size = Size(100, 25)
        self.open_button.Click += self.open_file_clicked

        self.Controls.Add(self.label)
        self.Controls.Add(self.open_button)

    def open_file_clicked(self, sender, args):
        try:
            os.startfile(self.filepath)
        except Exception as e:
            TaskDialog.Show("Error", "Failed to open file:\n{}".format(str(e)))
        self.Close()


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def parse_element_id(value):
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except Exception:
        try:
            return int(str(value).strip())
        except Exception:
            return None


def row_to_dict(headers, row):
    data = {}
    for i, header in enumerate(headers):
        key = str(header).strip()
        value = row[i] if i < len(row) else None
        data[key] = value
    return data


def sanitize_sheet_name(name):
    invalid = ['\\', '/', ':', '*', '?', '[', ']']
    safe = name or "Sheet"
    for ch in invalid:
        safe = safe.replace(ch, '_')
    safe = safe.strip()
    if not safe:
        safe = "Sheet"
    return safe[:31]


def get_level_map():
    return dict((lvl.Name, lvl) for lvl in DB.FilteredElementCollector(revit.doc).OfClass(DB.Level))


def get_field_parameter(element, field):
    param = None
    field_name = field.GetName()

    try:
        param = element.LookupParameter(field_name)
    except Exception:
        param = None

    if not param:
        try:
            pid = field.ParameterId
            if pid and pid != DB.ElementId.InvalidElementId:
                try:
                    bip = System.Enum.ToObject(DB.BuiltInParameter, get_eid_value(pid))
                    param = element.get_Parameter(bip)
                except Exception:
                    pass
        except Exception:
            pass

    return param


def get_parameter_export_value(param):
    if not param:
        return ""

    try:
        pname = ""
        try:
            pname = param.Definition.Name.lower()
        except Exception:
            pass

        if pname in ["level", "reference level", "schedule level"]:
            try:
                level_id = param.AsElementId()
                if level_id and get_eid_value(level_id) > 0:
                    level = revit.doc.GetElement(level_id)
                    return level.Name if level else ""
            except Exception:
                pass

        if param.StorageType == DB.StorageType.String:
            return param.AsString() or ""

        elif param.StorageType == DB.StorageType.Double:
            try:
                display = param.AsValueString()
                if display:
                    return display
            except Exception:
                pass
            try:
                return param.AsDouble()
            except Exception:
                return ""

        elif param.StorageType == DB.StorageType.Integer:
            try:
                return param.AsInteger()
            except Exception:
                return ""

        elif param.StorageType == DB.StorageType.ElementId:
            try:
                eid = param.AsElementId()
                if eid and get_eid_value(eid) > 0:
                    elem = revit.doc.GetElement(eid)
                    if elem and hasattr(elem, "Name"):
                        return elem.Name
                    return get_eid_value(eid)
            except Exception:
                return ""

    except Exception:
        return ""

    return ""


def set_parameter_import_value(param, value, level_map):
    if not param or param.IsReadOnly or value is None or value == "":
        return False

    try:
        pname = ""
        try:
            pname = param.Definition.Name.lower()
        except Exception:
            pass

        if pname in ["level", "reference level", "schedule level"]:
            level_name = str(value).strip()
            if level_name in level_map:
                param.Set(level_map[level_name].Id)
                return True
            return False

        if param.StorageType == DB.StorageType.String:
            param.Set(str(value))
            return True

        elif param.StorageType == DB.StorageType.Double:
            try:
                param.SetValueString(str(value))
                return True
            except Exception:
                pass
            try:
                param.Set(float(value))
                return True
            except Exception:
                return False

        elif param.StorageType == DB.StorageType.Integer:
            try:
                param.Set(int(float(value)))
                return True
            except Exception:
                return False

        elif param.StorageType == DB.StorageType.ElementId:
            try:
                eid_val = int(float(value))
                param.Set(get_elementid_from_value(eid_val))
                return True
            except Exception:
                return False

    except Exception:
        return False

    return False


# -----------------------------------------------------------------------------
# Export schedule to Excel
# -----------------------------------------------------------------------------

def export_schedule_to_excel(schedule, filepath):
    global last_export_dir

    try:
        schedule_view = schedule
        elements = list(DB.FilteredElementCollector(revit.doc, schedule_view.Id).ToElements())
        schedule_def = schedule_view.Definition
        fields = [schedule_def.GetField(fid) for fid in schedule_def.GetFieldOrder()]
        export_fields = [field for field in fields if field.ParameterId != DB.ElementId.InvalidElementId]

        headers = ["ElementId"] + [field.GetName() for field in export_fields]
        rows = [headers]

        for element in elements:
            row = [get_eid_value(element.Id)]

            for field in export_fields:
                param = get_field_parameter(element, field)
                value = get_parameter_export_value(param)
                row.append(value)

            rows.append(row)

        sheet_name = sanitize_sheet_name(schedule.Name)
        xl.dump(filepath, {sheet_name: rows})

        last_export_dir = os.path.dirname(filepath)

        open_dialog = OpenFileDialog(filepath)
        open_dialog.ShowDialog()

    except Exception as e:
        TaskDialog.Show("Error", "Export failed:\n{}".format(str(e).split('\n')[0]))



# -----------------------------------------------------------------------------
# Import changes from Excel
# -----------------------------------------------------------------------------

def import_changes_from_excel(filepath):
    if not os.path.exists(filepath):
        TaskDialog.Show("Error", "Excel file not found at:\n{}".format(filepath))
        return

    try:
        workbook_data = xl.load(filepath)
    except Exception as e:
        TaskDialog.Show("Error", "Failed to read Excel file:\n{}".format(str(e)))
        return

    if not workbook_data:
        TaskDialog.Show("Error", "No data found in Excel file.")
        return

    level_map = get_level_map()
    updated_count = 0
    rows_found = 0

    with revit.Transaction("Update Elements from Excel"):
        for sheet_name, sheet_data in workbook_data.items():
            headers = sheet_data.get("headers", [])
            rows = sheet_data.get("rows", [])

            if not headers or not rows:
                continue

            for row in rows:
                if isinstance(row, dict):
                    row_data = dict((str(k).strip(), v) for k, v in row.items())
                elif isinstance(row, list):
                    row_data = row_to_dict(headers, row)
                else:
                    continue

                elem_id_raw = row_data.get("ElementId")
                elem_id_val = parse_element_id(elem_id_raw)
                if elem_id_val is None:
                    continue

                element = revit.doc.GetElement(get_elementid_from_value(elem_id_val))
                if not element:
                    print "Element not found: {}".format(elem_id_val)
                    continue

                rows_found += 1

                for header, value in row_data.items():
                    if header == "ElementId":
                        continue

                    try:
                        param = element.LookupParameter(header)
                    except Exception:
                        param = None

                    if not param:
                        continue

                    if set_parameter_import_value(param, value, level_map):
                        updated_count += 1
                    else:
                        print "Skipped param '{}' for ElementId {}".format(header, elem_id_val)

    TaskDialog.Show(
        "Import Complete",
        "Rows found: {}\nParameters updated: {}".format(
            rows_found,
            updated_count
        )
    )


# -----------------------------------------------------------------------------
# Schedule picker dialog
# -----------------------------------------------------------------------------

class ScheduleDialog(Form):
    def __init__(self, schedules):
        self.schedules = schedules

        self.Text = "Schedule Import/Export"
        self.Width = 400
        self.Height = 300
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        self.search_label = Label()
        self.search_label.Text = "Search:"
        self.search_label.Location = Point(10, 5)
        self.search_label.Size = Size(50, 20)

        self.search_box = TextBox()
        self.search_box.Location = Point(70, 5)
        self.search_box.Size = Size(300, 20)
        self.search_box.TextChanged += self.search_changed

        self.schedule_listbox = ListBox()
        self.schedule_listbox.Location = Point(10, 30)
        self.schedule_listbox.Size = Size(360, 180)
        self.schedule_listbox.SelectionMode = SelectionMode.One

        for schedule in self.schedules:
            self.schedule_listbox.Items.Add(schedule.Name)

        self.export_button = Button()
        self.export_button.Text = "Export to Excel"
        self.export_button.Size = Size(100, 25)
        self.export_button.Location = Point((400 - 100 - 120 - 10) / 2, 230)
        self.export_button.Click += self.export_clicked

        self.import_button = Button()
        self.import_button.Text = "Import from Excel"
        self.import_button.Size = Size(120, 25)
        self.import_button.Location = Point(self.export_button.Location.X + self.export_button.Width + 15, 230)
        self.import_button.Click += self.import_clicked

        self.Controls.Add(self.search_label)
        self.Controls.Add(self.search_box)
        self.Controls.Add(self.schedule_listbox)
        self.Controls.Add(self.export_button)
        self.Controls.Add(self.import_button)

        self.selected_schedule = None
        self.action = None

    def search_changed(self, sender, args):
        search_term = self.search_box.Text.lower()
        self.schedule_listbox.Items.Clear()

        filtered_schedules = [s.Name for s in self.schedules if search_term in s.Name.lower()]
        self.schedule_listbox.Items.AddRange(Array[object](filtered_schedules))

    def export_clicked(self, sender, args):
        if self.schedule_listbox.SelectedIndex != -1:
            self.selected_schedule = self.schedule_listbox.SelectedItem
            self.action = "Export"
            self.Close()
        else:
            TaskDialog.Show("Error", "Please select a schedule to export.")

    def import_clicked(self, sender, args):
        self.action = "Import"
        self.Close()


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------

def main():
    global last_export_dir

    schedules = [
        s for s in DB.FilteredElementCollector(revit.doc)
        .OfClass(DB.ViewSchedule)
        .ToElements()
        if not s.IsTemplate
    ]

    if not schedules:
        TaskDialog.Show("Error", "No schedules found in the project.")
        return

    dialog = ScheduleDialog(schedules)
    dialog.ShowDialog()

    if not dialog.action:
        return

    if dialog.action == "Export":
        selected_schedule_name = dialog.selected_schedule
        selected_schedule = next(s for s in schedules if s.Name == selected_schedule_name)

        default_name = "{}.xlsx".format(selected_schedule.Name)
        filepath = forms.save_file(
            file_ext="xlsx",
            default_name=default_name,
            init_dir=last_export_dir,
            title="Save Schedule Export As",
            unc_paths=False
        )
        if filepath:
            if not filepath.lower().endswith(".xlsx"):
                filepath += ".xlsx"
            export_schedule_to_excel(selected_schedule, filepath)

    elif dialog.action == "Import":
        filepath = forms.pick_file(
            file_ext="xlsx",
            init_dir=last_export_dir,
            title="Select Excel File to Import"
        )
        if filepath:
            import_changes_from_excel(filepath)


if __name__ == "__main__":
    main()