import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, Transaction
from Autodesk.Revit.UI.Selection import PickBoxStyle

# Define the active Revit document and view
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Check if the active view is a Floor Plan (you can modify for other view types if needed)
if str(curview.ViewType) == 'FloorPlan':
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
    t = Transaction(doc, 'Set Crop Box for Active View')
    t.Start()
    
    # Set the crop box properties for the active view
    curview.CropBoxActive = True
    curview.CropBoxVisible = True
    curview.CropBox = bbox

    t.Commit()
else:
    print("Active view is not a Floor Plan view.")
