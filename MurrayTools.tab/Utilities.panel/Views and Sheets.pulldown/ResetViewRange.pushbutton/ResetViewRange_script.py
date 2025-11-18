import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI
import System

# Get active document and view
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
active_view = doc.ActiveView

# Show TaskDialog and exit quietly if not a floor plan
if active_view.ViewType != DB.ViewType.FloorPlan:
    dialog = UI.TaskDialog("Error")
    dialog.MainInstruction = "Script requires a floor plan view."
    dialog.Show()
    raise SystemExit

# Prompt user for view template modification
modify_view_template = False
if active_view.ViewTemplateId != DB.ElementId.InvalidElementId:
    dialog = UI.TaskDialog("Modify View Template")
    dialog.MainInstruction = "A view template is applied. Modify its view range?"
    dialog.CommonButtons = UI.TaskDialogCommonButtons.Yes | UI.TaskDialogCommonButtons.No
    dialog.DefaultButton = UI.TaskDialogResult.No
    result = dialog.Show()
    modify_view_template = (result == UI.TaskDialogResult.Yes)

# Get view range (from view template if chosen, else from view)
target_view = doc.GetElement(active_view.ViewTemplateId) if modify_view_template and active_view.ViewTemplateId != DB.ElementId.InvalidElementId else active_view
view_range = target_view.GetViewRange()

# Get associated level
associated_level_id = active_view.GenLevel.Id
if not associated_level_id:
    dialog = UI.TaskDialog("Error")
    dialog.MainInstruction = "No associated level found."
    dialog.Show()
    raise Exception("No associated level found.")

# Get levels sorted by elevation
levels = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
levels = sorted(levels, key=lambda x: x.Elevation)

# Find associated level and level above
associated_level = None
level_above = None
for level in levels:
    if level.Id == associated_level_id:
        associated_level = level
        for next_level in levels:
            if next_level.Elevation > level.Elevation:
                level_above = next_level
                break
        break

if not associated_level:
    dialog = UI.TaskDialog("Error")
    dialog.MainInstruction = "Associated level not found."
    dialog.Show()
    raise Exception("Associated level not found.")

# Start transaction
transaction = DB.Transaction(doc, "Reset View Range")
try:
    transaction.Start()

    # Set offsets
    top_offset = 0.0 if level_above else 4.0
    cut_offset = 4.0
    bottom_offset = 0.0
    view_depth_offset = 0.0

    # Set Top
    top_level = level_above if level_above else associated_level
    view_range.SetLevelId(DB.PlanViewPlane.TopClipPlane, top_level.Id)
    view_range.SetOffset(DB.PlanViewPlane.TopClipPlane, top_offset)

    # Set Cut Plane
    view_range.SetLevelId(DB.PlanViewPlane.CutPlane, associated_level.Id)
    view_range.SetOffset(DB.PlanViewPlane.CutPlane, cut_offset)

    # Set Bottom
    view_range.SetLevelId(DB.PlanViewPlane.BottomClipPlane, associated_level.Id)
    view_range.SetOffset(DB.PlanViewPlane.BottomClipPlane, bottom_offset)

    # Set View Depth
    view_range.SetLevelId(DB.PlanViewPlane.ViewDepthPlane, associated_level.Id)
    view_range.SetOffset(DB.PlanViewPlane.ViewDepthPlane, view_depth_offset)

    # Apply view range
    target_view.SetViewRange(view_range)

    transaction.Commit()
    
    # Show success dialog
    dialog = UI.TaskDialog("Success")
    dialog.MainInstruction = "View range reset successfully."
    dialog.Show()
except Exception as e:
    if transaction.HasStarted():
        transaction.RollBack()
    dialog = UI.TaskDialog("Error")
    dialog.MainInstruction = "Failed to set view range: {}".format(str(e))
    dialog.Show()
    raise Exception("Failed to set view range: {}".format(str(e)))