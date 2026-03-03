import os
from Autodesk.Revit.UI import TaskDialog

# File path (using raw string to handle backslashes)
file_path = r'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\MURRAY RIBBON\Murray.extension\MurrayTools.tab\Utilities.panel\ProjectUtilities.Stack\Exports.pulldown\OpenBOMTemplate.pushbutton\MR-BOM.xlsm'

if not os.path.exists(file_path):
    TaskDialog.Show("Warning", "File not Accessible: '{}' not found for schedule.".format(file_path))
else:
    # Open the file if path exists
    os.startfile(file_path)