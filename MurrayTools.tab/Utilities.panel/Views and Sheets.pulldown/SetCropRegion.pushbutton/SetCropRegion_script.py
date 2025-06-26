import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, Transaction, Transform, BuiltInParameter
from Autodesk.Revit.UI.Selection import PickBoxStyle
from Autodesk.Revit.UI import TaskDialog

# Define the active Revit document and view
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Check if the active view is a Floor Plan or Ceiling Plan
if str(curview.ViewType) in ['FloorPlan', 'CeilingPlan']:
    # Prompt the user to select a box area
    pickedBox = uidoc.Selection.PickBox(PickBoxStyle.Directional, "Select area for sketch")
    newmax = pickedBox.Max
    newmin = pickedBox.Min

    # Check if view is set to True North
    def is_true_north(view):
        orient_param = view.get_Parameter(BuiltInParameter.PLAN_VIEW_NORTH)
        return orient_param and orient_param.AsInteger() == 1  # 1 = True North

    # Adjust crop box based on view orientation
    def adjust_crop_box(view, max_pt, min_pt):
        # Initialize min and max points
        adj_max = XYZ(max_pt.X, max_pt.Y, 0)
        adj_min = XYZ(min_pt.X, min_pt.Y, 0)
        
        if is_true_north(view) and view.CropBox.Transform:
            # Transform view coordinates to model space using view's transform
            transform = view.CropBox.Transform.Inverse
            adj_max = transform.OfPoint(adj_max)
            adj_min = transform.OfPoint(adj_min)
        
        # Ensure min/max ordering
        corrected_min = XYZ(min(adj_min.X, adj_max.X), min(adj_min.Y, adj_max.Y), 0)
        corrected_max = XYZ(max(adj_min.X, adj_max.X), max(adj_min.Y, adj_max.Y), 0)
        
        # Create adjusted bounding box
        adjusted_bbox = BoundingBoxXYZ()
        adjusted_bbox.Min = corrected_min
        adjusted_bbox.Max = corrected_max
        adjusted_bbox.Transform = Transform.Identity
        
        return adjusted_bbox

    # Start a transaction to apply the changes
    t = Transaction(doc, 'Set Crop Box for Active View')
    t.Start()
    
    # Set the crop box properties for the active view
    adjusted_bbox = adjust_crop_box(curview, newmax, newmin)
    curview.CropBox = adjusted_bbox
    curview.CropBoxActive = True
    curview.CropBoxVisible = True
    t.Commit()
else:
    TaskDialog.Show("Error", "Active view is not a Floor or Ceiling Plan view.")