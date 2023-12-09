import os
from pyrevit import forms

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_WorksetMonitor.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    f = open((filepath), 'w')
    f.write('False')
    f.close()

f = open((filepath), 'r')
PrevInput = f.read()
f.close()

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


def toggle_subscription():
    from Autodesk.Revit.UI.Events import ViewActivatedEventArgs
    from System import EventHandler
    view_activated_handler = EventHandler[ViewActivatedEventArgs](on_view_activated)
    if PrevInput == 'True':
        __revit__.ViewActivated -= view_activated_handler
        f = open((filepath), 'w')
        f.write('False')
        f.close()
        forms.toast(
            'Disabled',
            title="Workset Monitor",
            appid="Murray Tools",
            icon="C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\MURRAY RIBBON\Murray.extension\Murray.ico",
            click="https://murraycompany.com",)
    else:
        __revit__.ViewActivated += view_activated_handler
        f = open((filepath), 'w')
        f.write('True')
        f.close()
        forms.toast(
            'Enabled',
            title="Workset Monitor",
            appid="Murray Tools",
            icon="C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\MURRAY RIBBON\Murray.extension\Murray.ico",
            click="https://murraycompany.com",)


# Set up the event handler and call the toggle_subscription function to subscribe/unsubscribe
toggle_subscription()
