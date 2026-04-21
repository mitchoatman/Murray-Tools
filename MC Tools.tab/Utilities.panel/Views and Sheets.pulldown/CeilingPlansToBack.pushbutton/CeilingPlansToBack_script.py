from collections import namedtuple
from pyrevit import revit, DB
from Autodesk.Revit.UI import TaskDialog

ViewportConfig = namedtuple('ViewportConfig', ['box_center',
                                               'type_id',
                                               'view_id',
                                               'label_offset'])

def capture_vp_config(viewport):
    return ViewportConfig(viewport.GetBoxCenter(),
                          viewport.GetTypeId(),
                          viewport.ViewId,
                          viewport.LabelOffset)

doc = revit.doc

# Collect all view sheets
sheets = DB.FilteredElementCollector(doc)\
           .OfClass(DB.ViewSheet)\
           .ToElements()

processed_sheets = 0

title = 'Ceiling Plans to Back'

with revit.Transaction('Put Ceiling Plans Behind Other Views'):
    for sheet in sheets:
        vp_ids = sheet.GetAllViewports()
        if len(vp_ids) == 0:
            continue

        vps = [doc.GetElement(vid) for vid in vp_ids]

        ceiling_vps = []
        non_ceiling_vps = []

        for vp in vps:
            view_id = vp.ViewId
            view = doc.GetElement(view_id) if view_id != DB.ElementId.InvalidElementId else None
            if view and view.ViewType == DB.ViewType.CeilingPlan:
                ceiling_vps.append(vp)
            else:
                non_ceiling_vps.append(vp)

        if len(ceiling_vps) == 0 or len(non_ceiling_vps) == 0:
            continue  # No action needed if only one type present

        # Sort non-ceiling viewports by ElementId to approximate original relative order
        non_ceiling_vps.sort(key=lambda vp: vp.Id.IntegerValue)

        # Delete and recreate non-ceiling viewports last (brings them to front)
        for vp in non_ceiling_vps:
            config = capture_vp_config(vp)
            sheet.DeleteViewport(vp)
            new_vp = DB.Viewport.Create(doc,
                                        sheet.Id,
                                        config.view_id,
                                        config.box_center)
            new_vp.ChangeTypeId(config.type_id)
            new_vp.LabelOffset = config.label_offset

        processed_sheets += 1

if processed_sheets > 0:
    message = 'Successfully reordered viewports on {} sheets.\nCeiling plans are now placed behind all other views.'.format(processed_sheets)
    TaskDialog.Show(title, message)
else:
    message = 'No sheets containing both ceiling plans and other views were found.'
    TaskDialog.Show(title, message)