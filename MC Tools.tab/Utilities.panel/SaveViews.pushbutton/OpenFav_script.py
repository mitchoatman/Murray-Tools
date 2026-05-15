from Autodesk.Revit.DB import View, ViewSchedule, ElementId
from Autodesk.Revit.UI import TaskDialog
import System
import os
import re

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)

# Fallback for unsaved files
if not file_name:
    file_name = doc.Title

folder_name = r"c:\Temp"
project_name = file_name.replace(" ", "_")
filepath = os.path.join(folder_name, 'Ribbon_OpenViews_{}.txt'.format(project_name))

def get_id_value(eid):
    try:
        return eid.Value        # Revit 2024+
    except:
        return eid.IntegerValue # older Revit

def make_element_id(val):
    try:
        return ElementId(System.Int64(val))  # Revit 2024+
    except:
        return ElementId(System.Int32(val))  # older Revit

if os.path.isfile(filepath):
    with open(filepath, 'r') as file:
        lines = [line.rstrip() for line in file.readlines()]

    if len(lines) < 2:
        TaskDialog.Show("Restore Views", "Saved views file is invalid.")
    else:
        saved_view_ids = []

        # Handles:
        # [12345, 67890]
        # [ElementId(12345), ElementId(67890)]
        # [[12345], [67890]]
        for s in lines[1][1:-1].split(','):
            s = s.strip()
            if not s:
                continue

            m = re.search(r'-?\d+', s)
            if m:
                saved_view_ids.append(int(m.group(0)))

        if lines[0] == str(file_name):
            for saved_id in saved_view_ids:
                view = doc.GetElement(make_element_id(saved_id))

                if not view or not isinstance(view, View):
                    continue

                if view.IsTemplate:
                    continue

                # Skip internal schedule views that Revit will not activate
                if isinstance(view, ViewSchedule):
                    if view.IsInternalKeynoteSchedule or view.IsTitleblockRevisionSchedule:
                        continue

                try:
                    uidoc.ActiveView = view
                except Exception as ex:
                    TaskDialog.Show(
                        "Restore Views",
                        "Could not open view '{}'\n{}".format(view.Name, ex)
                    )
        else:
            TaskDialog.Show("Invalid Views", "Saved views are not from this project")
else:
    TaskDialog.Show("Restore Views", "No Saved Views Found")