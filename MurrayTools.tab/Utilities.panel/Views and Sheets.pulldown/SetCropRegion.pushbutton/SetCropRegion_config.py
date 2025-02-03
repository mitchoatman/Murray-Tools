import Autodesk
from Autodesk.Revit.DB import BoundingBoxXYZ, XYZ, Transaction, Transform, ViewType, BuiltInParameter
from Autodesk.Revit.UI.Selection import PickBoxStyle
from pyrevit import forms

#---Define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

#---Get Selected Views
selected_views = forms.select_views(use_selection=True)
if not selected_views:
    forms.alert("No views selected. Please try again.", exitscript=True)

# Prompt the user to select a box area
pickedBox = uidoc.Selection.PickBox(PickBoxStyle.Directional, "Select area for sketch")
if not pickedBox:
    forms.alert("No area selected. Please try again.", exitscript=True)

# Get the max and min points of the selected box
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
selected_bbox = BoundingBoxXYZ()
selected_bbox.Max = XYZ(newmaxx, newmaxy, 0)
selected_bbox.Min = XYZ(newminx, newminy, 0)

# Function to check if a view is set to True North
def is_true_north(view):
    if view.ViewType in [ViewType.FloorPlan, ViewType.CeilingPlan]:
        orient_param = view.get_Parameter(BuiltInParameter.PLAN_VIEW_NORTH)
        if orient_param:
            return orient_param.AsInteger() == 1  # 1 = True North, 0 = Project North
    return False

# Function to adjust crop box based on view orientation
def adjust_crop_box(view, bbox):
    if is_true_north(view):
        # If the view is set to True North, transform the bounding box
        transform = view.CropBox.Transform
        if transform:
            inverted_transform = transform.Inverse
            adjusted_min = inverted_transform.OfPoint(bbox.Min)
            adjusted_max = inverted_transform.OfPoint(bbox.Max)
            
            # Ensure Min and Max points are correctly ordered
            corrected_min = XYZ(
                min(adjusted_min.X, adjusted_max.X),
                min(adjusted_min.Y, adjusted_max.Y),
                min(adjusted_min.Z, adjusted_max.Z)
            )
            corrected_max = XYZ(
                max(adjusted_min.X, adjusted_max.X),
                max(adjusted_min.Y, adjusted_max.Y),
                max(adjusted_min.Z, adjusted_max.Z)
            )
            
            # Create a valid bounding box
            adjusted_bbox = BoundingBoxXYZ()
            adjusted_bbox.Min = corrected_min
            adjusted_bbox.Max = corrected_max
            return adjusted_bbox
    return bbox  # Return original bounding box for Project North

# Start a transaction to apply the changes
t = Transaction(doc, 'Set Crop Box for Selected Views')
t.Start()

for view in selected_views:
    # Adjust the crop box based on the view orientation
    adjusted_bbox = adjust_crop_box(view, selected_bbox)
    # Set the crop box properties for the view
    view.CropBox = adjusted_bbox
    view.CropBoxActive = True
    view.CropBoxVisible = True

t.Commit()
