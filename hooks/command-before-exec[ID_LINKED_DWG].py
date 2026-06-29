# -*- coding: utf-8 -*-

from pyrevit import revit
from pyrevit.revit import events

TARGET_WORKSET_NAME = 'LINKS'


def move_links_to_links_workset(doc, target_name):
    try:
        from Autodesk.Revit import DB

        def get_target_workset(local_doc):
            try:
                for ws in DB.FilteredWorksetCollector(local_doc).OfKind(DB.WorksetKind.UserWorkset):
                    try:
                        if ws.Name and ws.Name.strip().upper() == target_name.upper():
                            return ws
                    except:
                        pass
            except:
                pass
            return None

        def is_link_element(elem):
            try:
                if isinstance(elem, DB.RevitLinkInstance):
                    return True

                if isinstance(elem, DB.ImportInstance):
                    try:
                        return elem.IsLinked
                    except:
                        return False
            except:
                pass

            return False

        if not doc or not doc.IsValidObject:
            return
        if doc.IsFamilyDocument:
            return
        if not doc.IsWorkshared:
            return

        target_ws = get_target_workset(doc)
        if not target_ws:
            return

        try:
            collector_rvt = list(
                DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
            )
        except:
            collector_rvt = []

        try:
            collector_cad = [
                x for x in DB.FilteredElementCollector(doc).OfClass(DB.ImportInstance)
                if is_link_element(x)
            ]
        except:
            collector_cad = []

        elements = collector_rvt + collector_cad
        if not elements:
            return

        target_id = target_ws.Id.IntegerValue
        changed = False
        t = DB.Transaction(doc, 'Move Links To LINKS Workset')

        try:
            t.Start()

            for elem in elements:
                try:
                    p = elem.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)

                    if not p or p.IsReadOnly:
                        continue

                    if p.AsInteger() == target_id:
                        continue

                    p.Set(target_id)
                    changed = True
                except:
                    pass

            if changed:
                t.Commit()
            else:
                t.RollBack()

        except:
            try:
                if t.HasStarted():
                    t.RollBack()
            except:
                pass

    except:
        pass


# =========================================================
# HOOK ENTRY
# =========================================================

try:
    doc = revit.doc

    if doc:
        events.execute_in_revit_context(
            move_links_to_links_workset,
            doc,
            TARGET_WORKSET_NAME
        )
except:
    pass