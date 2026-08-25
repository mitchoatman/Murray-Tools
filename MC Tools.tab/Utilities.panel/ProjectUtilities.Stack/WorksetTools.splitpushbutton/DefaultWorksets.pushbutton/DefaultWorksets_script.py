# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import (
    Workset, WorksetTable, Transaction,
    FilteredWorksetCollector, WorksetKind
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommandLinkId, TaskDialogCommonButtons, TaskDialogResult

doc = __revit__.ActiveUIDocument.Document

TARGET_LEVELS_WS = "MURRAY Levels and Grids"
OTHER_DEFAULT_LEVEL_NAMES = [
    "Shared Views, Levels, Grids",
    "Shared Levels and Grids"
]

WORKSETS_TO_ADD = [
    "LINKS",
    "POINT_LAYOUT"
]


def get_user_worksets(document):
    return list(FilteredWorksetCollector(document).OfKind(WorksetKind.UserWorkset))


def get_workset_by_name(document, name):
    for ws in get_user_worksets(document):
        if ws.Name == name:
            return ws
    return None


def find_default_levels_workset(document):
    for name in OTHER_DEFAULT_LEVEL_NAMES:
        ws = get_workset_by_name(document, name)
        if ws:
            return ws
    return None


def prompt_for_trade():
    td = TaskDialog("Select Trade")
    td.MainInstruction = "Choose trade for Workset1"
    td.MainContent = "The selected trade will be used to rename Workset1 or as the default workset when enabling worksharing."
    td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "MECH_PIPE")
    td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "MECH_DUCT")
    td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "PLUMBING")
    td.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "PROCESS")
    td.CommonButtons = TaskDialogCommonButtons.Cancel

    result = td.Show()

    if result == TaskDialogResult.CommandLink1:
        return "MECH_PIPE"
    elif result == TaskDialogResult.CommandLink2:
        return "MECH_DUCT"
    elif result == TaskDialogResult.CommandLink3:
        return "PLUMBING"
    elif result == TaskDialogResult.CommandLink4:
        return "PROCESS"
    else:
        return None


selected_trade = prompt_for_trade()

if not selected_trade:
    TaskDialog.Show("Cancelled", "No trade selected. Script cancelled.")
else:
    # Enable worksharing if needed
    if not doc.IsWorkshared:
        try:
            if doc.IsModelInCloud:
                if doc.CanEnableCloudWorksharing():
                    doc.EnableCloudWorksharing()
                else:
                    raise Exception("Cloud worksharing cannot be enabled for this model.")
            elif doc.CanEnableWorksharing():
                doc.EnableWorksharing(TARGET_LEVELS_WS, selected_trade)
            else:
                raise Exception("Worksharing cannot be enabled for this model.")
        except Exception as e:
            TaskDialog.Show("Error", "Error enabling worksharing:\n\n{}".format(str(e)))

    if not doc.IsWorkshared:
        TaskDialog.Show("Error", "Worksharing is not enabled. Cannot configure worksets.")
    else:
        t = Transaction(doc, "Configure Murray Worksets")
        t.Start()
        try:
            # Rename default levels/grids workset if needed
            default_levels_ws = find_default_levels_workset(doc)
            target_levels_ws = get_workset_by_name(doc, TARGET_LEVELS_WS)

            if default_levels_ws and not target_levels_ws:
                WorksetTable.RenameWorkset(doc, default_levels_ws.Id, TARGET_LEVELS_WS)

            # Rename Workset1 to selected trade if possible
            default_model_ws = get_workset_by_name(doc, "Workset1")
            selected_trade_ws = get_workset_by_name(doc, selected_trade)

            if default_model_ws and not selected_trade_ws:
                WorksetTable.RenameWorkset(doc, default_model_ws.Id, selected_trade)

            # Add LINKS and POINT_LAYOUT only
            existing_names = [ws.Name for ws in get_user_worksets(doc)]
            for ws_name in WORKSETS_TO_ADD:
                if ws_name not in existing_names:
                    Workset.Create(doc, ws_name)

            t.Commit()
            TaskDialog.Show("Success", "Worksets configured successfully.")
        except Exception as e:
            t.RollBack()
            TaskDialog.Show("Error", "Error configuring worksets:\n\n{}".format(str(e)))