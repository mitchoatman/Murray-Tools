# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (BoundingBoxXYZ, FilteredElementCollector, Transaction,
                               BuiltInCategory, FabricationPart, Level, LogicalOrFilter,
                               ElementCategoryFilter, ElementId, WorksharingUtils)
from System.Collections.Generic import List
import System

doc = __revit__.ActiveUIDocument.Document
curview = doc.ActiveView

# Collect all Level elements (in active view), storing their names, elevations, and IDs
level_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType()
level_elevations = {}
level_ids = {}
for level in level_collector:
    level_elevations[level.Name] = level.Elevation
    level_ids[level.Name] = level.Id

# Sort levels by elevation (ascending order)
sorted_levels = sorted(level_elevations.items(), key=lambda x: x[1])

if not sorted_levels:
    raise Exception("No levels visible in the active view.")

# Function to get the bottom Z elevation of a fabrication part's bounding box
def get_bottom_point_z(ele):
    bBox = ele.get_BoundingBox(None)
    if bBox is None:
        return None
    return bBox.Min.Z

# Function to get the closest level at or below z (lowest level if z is below all)
def get_level_id_for_z(z):
    assigned_level_id = level_ids[sorted_levels[0][0]]
    for level_name, elev in sorted_levels:
        if elev <= z:
            assigned_level_id = level_ids[level_name]
        else:
            break
    return assigned_level_id

# Category filters for Fabrication Pipework and Fabrication Ductwork
pipework_filter = ElementCategoryFilter(BuiltInCategory.OST_FabricationPipework)
ductwork_filter = ElementCategoryFilter(BuiltInCategory.OST_FabricationDuctwork)
combined_filter = LogicalOrFilter(pipework_filter, ductwork_filter)

# Collect all MEP Fabrication Pipework and Ductwork elements in the current view
fabrication_elements = FilteredElementCollector(doc, curview.Id) \
                       .OfClass(FabricationPart) \
                       .WherePasses(combined_filter) \
                       .WhereElementIsNotElementType() \
                       .ToElements()

# Pass 1: determine which parts actually need a change (no edits yet)
changes = []  # list of (element, new_level_id)
for elem in fabrication_elements:
    bottom_z = get_bottom_point_z(elem)
    if bottom_z is None:
        continue
    param = elem.LookupParameter("Reference Level")
    if param is None or param.IsReadOnly:
        continue
    new_level_id = get_level_id_for_z(bottom_z)
    if param.AsElementId() != new_level_id:
        changes.append((elem, new_level_id))

# Pass 2: check out everything in ONE call, skip what is owned by someone else
skipped = set()
if doc.IsWorkshared and changes:
    ids = List[ElementId]([e.Id for e, _ in changes])
    failed = WorksharingUtils.CheckoutElements(doc, ids)
    skipped = set(i.IntegerValue for i in failed)

# Pass 3: apply changes in a single transaction
t = Transaction(doc, "Assign Levels to Fabrication Parts")
t.Start()
done = 0
for elem, new_level_id in changes:
    if elem.Id.IntegerValue in skipped:
        continue
    elem.LookupParameter("Reference Level").Set(new_level_id)
    done += 1
t.Commit()

# print("Updated: {}  |  Unchanged: {}  |  Skipped (owned by another user): {}".format(
    # done, len(fabrication_elements) - len(changes), len(skipped)))