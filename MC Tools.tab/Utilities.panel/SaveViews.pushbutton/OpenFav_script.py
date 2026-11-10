from Autodesk.Revit.DB import View
from Autodesk.Revit.UI import TaskDialog
import System
import os
import sys
import json

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

folder_name = r"c:\Temp"

def get_id_value(eid):
    try:
        return eid.Value        # Revit 2024+
    except:
        return eid.IntegerValue # older versions

# Use the same project identifier logic as restore
file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)

# If unsaved model, fall back to title
if not file_name:
    file_name = doc.Title

open_views = []
for uiview in uidoc.GetOpenUIViews():
    view = doc.GetElement(uiview.ViewId)
    if view and not view.IsTemplate:
        open_views.append(view)

if not open_views:
    TaskDialog.Show("Warning", "There are no open views.")
    sys.exit()

if len(open_views) > 10:
    TaskDialog.Show("Warning", "You have more than ten open views. Opening this many views at once may take some time.")

view_list = [get_id_value(view.Id) for view in open_views]

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

project_name = file_name.replace(" ", "_")
filepath = os.path.join(folder_name, "Ribbon_OpenViews_{}.json".format(project_name))

# Structure payload for JSON
data = {
    "project_name": file_name,
    "view_count": len(view_list),
    "view_ids": view_list
}

with open(filepath, 'w') as the_file:
    json.dump(data, the_file, indent=4)

TaskDialog.Show("Status", "[{}] views have been saved.".format(len(view_list)))