import os
from Autodesk.Revit.UI import TaskDialog

# File path (using raw string to handle backslashes)
file_path = r"Z:\shared\Reference\!Detailing\12-Cad Request\Super Secret Folder\!CAD Group Request_MASTER v2.5.xlsm"

if not os.path.exists(file_path):
    TaskDialog.Show("Warning", "File not Accessible: '{}' not found for schedule.".format(file_path))
else:
    # Open the file if path exists
    os.startfile(file_path)