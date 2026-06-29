# -*- coding: utf-8 -*-

from pyrevit import revit
from pyrevit.revit import events

TARGET_WORKSET_NAME = 'LINKS'


def move_links_to_links_workset(doc, target_name):
    try:
        from Autodesk.Revit import DB

        # Safety checks
        if not doc or not doc.IsValidObject:
            return
        if doc.IsFamilyDocument:
            return
        if not doc.IsWorkshared:
            return

        # Find target workset quietly
        target_ws = None
        for ws in DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset):
            try:
                if ws.Name and ws.Name.strip().upper() == target_name.upper():
                    target_ws = ws
                    break
            except:
                pass

        # Older models may not have the workset
        if not target_ws:
            return

        # Collect Revit links
        elements = []
        try:
            elements.extend(DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance))
        except:
            pass

        # Collect linked CAD imports only
        try:
            for elem in DB.FilteredElementCollector(doc).OfClass(DB.ImportInstance):
                try:
                    if elem.IsLinked:
                        elements.append(elem)
                except:
                    pass
        except:
            pass

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