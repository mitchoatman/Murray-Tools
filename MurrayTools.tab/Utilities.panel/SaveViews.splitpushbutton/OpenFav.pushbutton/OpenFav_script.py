from pyrevit import forms
import Autodesk
from Autodesk.Revit.DB import *
import System
import os
import re

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)

folder_name = "c:\\Temp"

# Replace spaces in the project name with underscores
project_name = file_name.replace(" ", "_")

# Dynamically generate the file name with the project name
filepath = os.path.join(folder_name, 'Ribbon_OpenViews_{}.txt'.format(project_name))

AllViews = FilteredElementCollector(doc).OfClass(View)
AllViewNames = [view.Name for view in AllViews]

# Check if the file with the dynamically generated name exists
if os.path.isfile(filepath):
    with open(filepath, 'r') as file:
        lines = file.readlines()
        lines = [line.rstrip() for line in lines]

    # Extract the IDs from the strings
    saved_view_ids = [int(re.search(r'\[(\d+)\]', line).group(1)) for line in lines[1][1:-1].split(', ')]

    if lines[0] == str(file_name):
        for view in AllViews:
            # Check if the view's Id is in the saved list
            if view.Id.IntegerValue in saved_view_ids:
                if not view.IsTemplate and view.CanBePrinted:  # Check if the view is not a template and can be printed
                    ViewToOpen = doc.GetElement(view.Id)  # Use the Id property of the view
                    uidoc.RequestViewChange(ViewToOpen)
    else:
        print('Saved views are not from this project')
else:
    print 'No Saved Views Found'
