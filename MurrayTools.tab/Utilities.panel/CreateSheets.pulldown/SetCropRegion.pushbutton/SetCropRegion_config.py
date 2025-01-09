import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, Transaction, ViewDuplicateOption
from Autodesk.Revit.UI.Selection import PickBoxStyle
from pyrevit import forms
import sys

#---Define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application

#---Get Selected Views
selected_views = forms.select_views(use_selection=True)
if not selected_views:
    forms.alert("No views selected. Please try again.", exitscript = True)



# Check if the active view is a Floor Plan (you can modify for other view types if needed)
# if str(curview.ViewType) == 'FloorPlan':
# Prompt the user to select a box area
pickedBox = uidoc.Selection.PickBox(PickBoxStyle.Directional, "Select area for sketch")
Maxx = pickedBox.Max.X
Maxy = pickedBox.Max.Y
Minx = pickedBox.Min.X
Miny = pickedBox.Min.Y

# Ensure min and max coordinates are correctly ordered
newmaxx = max(Maxx, Minx)
newmaxy = max(Maxy, Miny)
newminx = min(Maxx, Minx)
newminy = min(Maxy, Miny)

# Create a bounding box based on the selected area
bbox = BoundingBoxXYZ()
bbox.Max = XYZ(newmaxx, newmaxy, 0)
bbox.Min = XYZ(newminx, newminy, 0)

# Start a transaction to apply the changes
t = Transaction(doc, 'Set Crop Box for Selected Views')
t.Start()

for view in selected_views:
    # Set the crop box properties for the active view
    view.CropBoxActive = True
    view.CropBoxVisible = True
    view.CropBox = bbox

t.Commit()
# else:
    # print("Active view is not a Floor Plan view.")
