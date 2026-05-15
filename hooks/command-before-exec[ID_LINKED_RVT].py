# -*- coding: utf-8 -*-

from pyrevit import revit
from pyrevit.revit import events
TARGET_WORKSET_NAME = 'LINKS'


def move_links_to_links_workset(doc, target_name):

    import Autodesk
    from Autodesk.Revit import DB

    from pyrevit import script

    logger = script.get_logger()

    def get_target_workset(local_doc):
        for ws in DB.FilteredWorksetCollector(local_doc)\
                    .OfKind(DB.WorksetKind.UserWorkset):
            if ws.Name.strip().upper() == target_name.upper():
                return ws
        return None


    def is_link_element(elem):
        if isinstance(elem, DB.RevitLinkInstance):
            return True

        if isinstance(elem, DB.ImportInstance):
            try:
                return elem.IsLinked
            except:
                return False

        return False


    if not doc or not doc.IsValidObject or not doc.IsWorkshared:
        return

    target_ws = get_target_workset(doc)
    if not target_ws:
        return

    collector_rvt = DB.FilteredElementCollector(doc)\
        .OfClass(DB.RevitLinkInstance)

    collector_cad = [
        x for x in DB.FilteredElementCollector(doc)
        .OfClass(DB.ImportInstance)
        if is_link_element(x)
    ]

    elements = list(collector_rvt) + collector_cad

    t = DB.Transaction(doc, 'Move Links To LINKS Workset')

    moved = 0

    try:
        t.Start()

        for elem in elements:

            p = elem.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)

            if not p or p.IsReadOnly:
                continue

            if p.AsInteger() == target_ws.Id.IntegerValue:
                continue

            p.Set(target_ws.Id.IntegerValue)
            moved += 1

        t.Commit()

        if moved:
            logger.info(
                'Moved %s link(s) to workset "%s".',
                moved,
                target_name
            )

    except Exception as ex:
        try:
            if t.HasStarted():
                t.RollBack()
        except:
            pass

        logger.error(str(ex))


# =========================================================
# HOOK ENTRY
# =========================================================

doc = revit.doc

if doc:
    events.execute_in_revit_context(
        move_links_to_links_workset,
        doc,
        TARGET_WORKSET_NAME   # <-- PASS IT HERE
    )