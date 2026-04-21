# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import ElementId, XYZ, BoundingBoxXYZ, Outline, Transform
from Autodesk.Revit.UI import UIApplication, TaskDialog
from System.Collections.Generic import List
import math
# Get the active Revit document and UI document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
# Get the currently selected elements
selected_ids = uidoc.Selection.GetElementIds()
# Check if there is at least one selected element
if selected_ids.Count > 0:
    elements = [doc.GetElement(eid) for eid in selected_ids]
    curview = doc.ActiveView
    valid_bbs = []
    for elem in elements:
        bb = elem.get_BoundingBox(curview)
        if bb and bb.Min.X != bb.Min.X:  # Check for NaN (x != x if NaN)
            continue
        if bb:
            valid_bbs.append(bb)
    if valid_bbs:
        # Initialize union with first BB
        union_min = valid_bbs[0].Min
        union_max = valid_bbs[0].Max
        # Union with remaining BBs
        for bb in valid_bbs[1:]:
            union_min = XYZ(
                min(union_min.X, bb.Min.X),
                min(union_min.Y, bb.Min.Y),
                min(union_min.Z, bb.Min.Z)
            )
            union_max = XYZ(
                max(union_max.X, bb.Max.X),
                max(union_max.Y, bb.Max.Y),
                max(union_max.Z, bb.Max.Z)
            )
        # Ensure finite values (fallback if any inf/NaN slipped through)
        coords = [union_min.X, union_min.Y, union_min.Z, union_max.X, union_max.Y, union_max.Z]
        if any(math.isinf(c) or math.isnan(c) for c in coords):
            TaskDialog.Show("Invalid Bounds", "Selected elements have invalid bounding boxes (infinite or NaN values).")
        else:
            from Autodesk.Revit.DB import Transaction
            t = Transaction(doc, 'Temporary Crop for Zoom')
            t.Start()
            try:
                original_crop = curview.CropBox
                original_active = curview.CropBoxActive
                outline = Outline(union_min, union_max)
                tform = Transform.Identity
                new_crop = BoundingBoxXYZ.Create(tform, outline)
                curview.CropBox = new_crop
                curview.CropBoxActive = True
                t.Commit()
                uidoc.ZoomToFit()
                # Restore original crop settings
                t_restore = Transaction(doc, 'Restore Crop Box')
                t_restore.Start()
                curview.CropBoxActive = original_active
                if original_crop:
                    curview.CropBox = original_crop
                t_restore.Commit()
            except Exception as ex:
                t.RollBack()
                # Fallback to ShowElements
                element_id_list = List[ElementId](selected_ids)
                uidoc.ShowElements(element_id_list)
    else:
        TaskDialog.Show("No Visible Elements", "Selected elements are not visible or have invalid bounding boxes in the active view.")
else:
    # Show message if no elements are selected
    TaskDialog.Show("No Selection", "Please select at least one element to zoom to.")