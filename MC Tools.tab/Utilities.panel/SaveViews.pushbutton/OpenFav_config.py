from Autodesk.Revit.DB import View
from Autodesk.Revit.UI import TaskDialog
import System
import os
import sys

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

project_name = file_name.replace(" ", "_")
filepath = os.path.join(folder_name, "Ribbon_OpenViews_{}.txt".format(project_name))

with open(filepath, 'w') as the_file:
    the_file.write(str(file_name) + '\n')
    the_file.write(str(view_list) + '\n')

TaskDialog.Show("Status", "[{}] views have been saved.".format(len(view_list)))