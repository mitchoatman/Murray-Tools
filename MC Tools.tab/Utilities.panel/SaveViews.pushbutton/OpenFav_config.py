from Autodesk.Revit.DB import FilteredElementCollector, View
from Autodesk.Revit.UI import TaskDialog
import System
import os
import re

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)

folder_name = "c:\\Temp"

file_name = doc.Title

open_views = [doc.GetElement(view.ViewId) for view in uidoc.GetOpenUIViews() if not doc.GetElement(view.ViewId).IsTemplate]

if not open_views:
    TaskDialog.Show("Warning", 'There are no open views.')
    sys.exit()

if len(open_views) > 10:
    TaskDialog.Show("Warning", "You have more than ten open views, Opening this many open views at once may take some time.")


view_list = [view.Id for view in open_views]

folder_name = "c:\\Temp"

# Replace spaces in the project name with underscores
project_name = file_name.replace(" ", "_")

# Append the project name to the file path using format method
filepath = os.path.join(folder_name, 'Ribbon_OpenViews_{}.txt'.format(project_name))

# Write values to a text file for future retrieval
with open(filepath, 'w') as the_file:
    line1 = str(file_name) + '\n'
    line2 = str(view_list) + '\n'
    the_file.writelines([line1, line2])

TaskDialog.Show("Status", "[{}] views have been saved.".format(len(view_list)))


