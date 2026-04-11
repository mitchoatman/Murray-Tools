from Autodesk.Revit.DB import FilteredElementCollector, View, ViewSchedule, ElementId
from Autodesk.Revit.UI import TaskDialog
import System, os, re

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application

file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)

folder_name = r"c:\Temp"
project_name = file_name.replace(" ", "_")
filepath = os.path.join(folder_name, 'Ribbon_OpenViews_{}.txt'.format(project_name))

def get_id_value(eid):
    try:
        return eid.Value        # Revit 2026+
    except:
        return eid.IntegerValue # older versions

if os.path.isfile(filepath):
    with open(filepath, 'r') as file:
        lines = [line.rstrip() for line in file.readlines()]

    saved_view_ids = [int(re.search(r'\[(\d+)\]', s).group(1))
                      for s in lines[1][1:-1].split(', ')]

    if lines[0] == str(file_name):
        for saved_id in saved_view_ids:
            view = doc.GetElement(ElementId(saved_id))
            if not view or not isinstance(view, View):
                continue
            if view.IsTemplate:
                continue

            # Skip internal schedule views that Revit won't activate
            if isinstance(view, ViewSchedule):
                if view.IsInternalKeynoteSchedule or view.IsTitleblockRevisionSchedule:
                    continue

            try:
                uidoc.ActiveView = view
            except Exception as ex:
                TaskDialog.Show("Restore Views",
                                "Could not open view '{}'\n{}".format(view.Name, ex))
    else:
        TaskDialog.Show("Invalid Views", "Saved views are not from this project")
else:
    TaskDialog.Show("Restore Views", "No Saved Views Found")