from Autodesk.Revit.DB import View, ViewSchedule, ElementId
from Autodesk.Revit.UI import TaskDialog
import System
import os
import json

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)

# Fallback for unsaved files
if not file_name:
    file_name = doc.Title

folder_name = r"c:\Temp"
project_name = file_name.replace(" ", "_")
filepath = os.path.join(folder_name, 'Ribbon_OpenViews_{}.json'.format(project_name))

def make_element_id(val):
    try:
        return ElementId(System.Int64(val))  # Revit 2024+
    except:
        return ElementId(System.Int32(val))  # older Revit

if os.path.isfile(filepath):
    try:
        with open(filepath, 'r') as file:
            data = json.load(file)
        
        saved_project_name = data.get("project_name")
        saved_view_ids = data.get("view_ids", [])

        if saved_project_name == str(file_name):
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

    except (ValueError, KeyError, json.JSONDecodeError) as err:
        TaskDialog.Show("Restore Views", "Saved views file is invalid or corrupted.\n{}".format(err))
else:
    TaskDialog.Show("Restore Views", "No Saved Views Found")