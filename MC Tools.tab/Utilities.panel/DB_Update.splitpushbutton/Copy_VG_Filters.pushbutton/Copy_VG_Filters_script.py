import clr
import os
import os.path as op
import pickle as pl
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from pyrevit import script, revit
from pyrevit.revit import doc, uidoc, selection
from Autodesk.Revit.UI import UIDocument, TaskDialog

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

output = script.get_output()
logger = script.get_logger()
linkify = output.linkify
selection = selection.get_selection()
selected_ids = selection.element_ids

ALLOWED_TYPES = [
    ViewPlan,
    View3D,
    ViewSection,
    ViewSheet,
    ViewDrafting
]

def get_view_templates(doc):
    allview_templates = FilteredElementCollector(doc)\
        .WhereElementIsNotElementType()\
        .OfCategory(BuiltInCategory.OST_Views)\
        .ToElements()
    viewtemplate_list = []
    for vt in allview_templates:
        if vt.IsTemplate:
            viewtemplate_list.append(vt)
    return viewtemplate_list

def get_view_filters(doc, view):
    result = []
    for filter_id in view.GetFilters():
        filter_element = doc.GetElement(filter_id)
        result.append(filter_element)
    return result

def get_active_view():
    if isinstance(doc.ActiveView, View):
        active_view = doc.ActiveView
    else:
        if len(selected_ids) == 0:
            logger.error('Select a view with applied template, to copy filters from it')
            return None
        active_view = doc.GetElement(selected_ids[0])
        if not isinstance(active_view, ALLOWED_TYPES):
            logger.error('Selected view is not allowed. Please select or open view from which '
                         'you want to copy template settings VG Overrides - Filters')
            return None
    return active_view

def main():
    active_view = get_active_view()
    if not active_view:
        logger.warning('Activate a view to copy')
        return

    logger.debug('Source view selected: %s id%s' % (active_view.Name, active_view.Id.ToString()))

    # Get all filters currently applied to the source view/template
    active_filters = get_view_filters(doc, active_view)
    if not active_filters:
        logger.warning('Active view has no filter overrides')
        return

    # Get all view templates in the project
    viewtemplate_list = get_view_templates(doc)
    if not viewtemplate_list:
        logger.warning('Project has no view templates')
        return

    logger.info('Found %d view templates in project' % len(viewtemplate_list))
    logger.info('Copying %d filters to all templates' % len(active_filters))

    t = Transaction(doc, "Copy All VG Filters to All Templates")
    t.Start()

    for vt in viewtemplate_list:
        # Skip the source itself if it is a template
        if vt.Id == active_view.Id:
            continue

        for filter_element in active_filters:
            fid = filter_element.Id
            try:
                vt.RemoveFilter(fid)
                logger.debug('filter %s deleted from template %s' % (fid.IntegerValue.ToString(), vt.Name))
            except:
                pass

            try:
                fr = active_view.GetFilterOverrides(fid)
                vt.SetFilterOverrides(fid, fr)
            except Exception as e:
                logger.warning('filter %s was not applied to template %s - %s' % (
                    fid.IntegerValue.ToString(), vt.Name, str(e)))

    t.Commit()

    view_count = len(viewtemplate_list)
    TaskDialog.Show(
        "Copy Filters",
        "Filters copied to all view templates.\n\n"
      + "**All filters from the active view/template have been copied to {} view templates.**".format(view_count)
    )

if __name__ == "__main__":
    main()