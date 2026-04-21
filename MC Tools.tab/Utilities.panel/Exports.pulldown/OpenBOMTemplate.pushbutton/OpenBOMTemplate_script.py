import os
from Autodesk.Revit.UI import TaskDialog

path, filename = os.path.split(__file__)
Filename = '\MR-BOM.xlsm'

# File path (using raw string to handle backslashes)
file_path = path + Filename

if not os.path.exists(file_path):
    TaskDialog.Show("Warning", "File not Accessible: '{}' not found for schedule.".format(file_path))
else:
    # Open the file if path exists
    os.startfile(file_path)