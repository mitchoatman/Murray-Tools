from Autodesk.Revit.DB import (
    Workset, WorksetTable, Transaction,
    FilteredWorksetCollector, WorksetKind
)
from Autodesk.Revit.UI import TaskDialog

doc = __revit__.ActiveUIDocument.Document

TARGET_LEVELS_WS = 'MURRAY Levels and Grids'
OTHER_DEFAULT_NAMES = [
    'Shared Views, Levels, Grids',
    'Shared Levels and Grids'
]

WORKSETS_TO_ADD = [
    'LINKS',
    'POINTLAYOUT'
]


def get_user_worksets(document):
    return list(FilteredWorksetCollector(document).OfKind(WorksetKind.UserWorkset))


def get_workset_by_name(document, name):
    for ws in get_user_worksets(document):
        if ws.Name == name:
            return ws
    return None


def find_default_levels_workset(document):
    for name in OTHER_DEFAULT_NAMES:
        ws = get_workset_by_name(document, name)
        if ws:
            return ws
    return None


# 1) Enable worksharing if needed
if not doc.IsWorkshared:
    try:
        if doc.IsModelInCloud:
            if doc.CanEnableCloudWorksharing():
                doc.EnableCloudWorksharing()
            else:
                raise Exception("Cloud worksharing cannot be enabled for this model.")
        elif doc.CanEnableWorksharing():
            # Correct parameter order:
            # (levels/grids, all other model elements)
            doc.EnableWorksharing(TARGET_LEVELS_WS, 'Workset1')
        else:
            raise Exception("Worksharing cannot be enabled for this model.")
    except Exception as e:
        TaskDialog.Show("Error", "Error enabling worksharing:\n\n{}".format(str(e)))

if not doc.IsWorkshared:
    TaskDialog.Show("Error", "Worksharing is not enabled. Cannot create worksets.")
else:
    t = Transaction(doc, "Configure Murray Worksets")
    t.Start()
    try:
        # For cloud models, rename the default levels/grids workset first
        default_ws = find_default_levels_workset(doc)
        target_ws = get_workset_by_name(doc, TARGET_LEVELS_WS)

        if default_ws and not target_ws:
            WorksetTable.RenameWorkset(doc, default_ws.Id, TARGET_LEVELS_WS)

        # Add only the remaining worksets
        existing_names = [ws.Name for ws in get_user_worksets(doc)]
        for ws_name in WORKSETS_TO_ADD:
            if ws_name not in existing_names:
                Workset.Create(doc, ws_name)

        t.Commit()
        TaskDialog.Show("Success", "Murray worksets configured successfully.")
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "Error configuring worksets:\n\n{}".format(str(e)))