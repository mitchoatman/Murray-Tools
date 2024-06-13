import os
from pyrevit import forms
from Autodesk.Revit.UI.Events import ViewActivatedEventArgs
from System import EventHandler


def on_view_activated(sender, args):
    from Autodesk.Revit.DB import FilteredWorksetCollector
    active_doc = sender.ActiveUIDocument.Document
    active_view = active_doc.ActiveView
    WorksetNames = []
    WorksetId = []
    AllWorksets = FilteredWorksetCollector(active_doc)
    for c in AllWorksets:
        WorksetNames.append(c.Name)
        WorksetId.append(c.Id)
    if str(active_view.ViewType) == 'FloorPlan':
        # Get the associated level of the active view
        level = active_view.GenLevel
        # Searches worksetnames for level name and returns index
        index = [i for i, name in enumerate(WorksetNames) if name == level.Name][0]
        # Sets the active workset to match active view associated level
        active_doc.GetWorksetTable().SetActiveWorksetId(WorksetId[index])

view_activated_handler = EventHandler[ViewActivatedEventArgs](on_view_activated)
__revit__.ViewActivated -= view_activated_handler


