# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.UI import UIApplication
from System.Collections.Generic import List

# Get the active Revit document and UI document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Get the currently selected elements
selected_ids = uidoc.Selection.GetElementIds()

# Check if there is at least one selected element
if selected_ids.Count > 0:
    # Convert to List[ElementId] for ShowElements
    element_id_list = List[ElementId](selected_ids)
    # Zoom to the selected elements in the active view
    uidoc.ShowElements(element_id_list)
else:
    # Show message if no elements are selected
    from Autodesk.Revit.UI import TaskDialog
    TaskDialog.Show("No Selection", "Please select at least one element to zoom to.")