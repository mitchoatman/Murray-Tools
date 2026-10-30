# -*- coding: utf-8 -*-
import sys
import traceback
from Autodesk.Revit.DB import (Transaction, TransactionStatus, FilteredElementCollector,
                               BuiltInCategory, FamilyInstance, IndependentTag, TextNote,
                               Dimension, ElementId, WorksharingUtils, CheckoutStatus)
from System.Collections.Generic import List
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

PRECISION = 4  # decimal places (feet) used when comparing positions


def eid(i):
    """Integer value of an ElementId (older and newer Revit versions)."""
    try:
        return i.Value
    except AttributeError:
        return i.IntegerValue


def bic(name):
    return getattr(BuiltInCategory, name, None)


MODEL_CATEGORY_NAMES = {
    "Structural Framing": "OST_StructuralFraming",
    "Structural Columns": "OST_StructuralColumns",
    "Structural Stiffener": "OST_StructuralStiffener",
    "Pipe Accessory": "OST_PipeAccessory",
    "Plumbing Fixtures": "OST_PlumbingFixtures",
    "Mechanical Equipment": "OST_MechanicalEquipment",
    "Duct Accessory": "OST_DuctAccessory",
    "Generic Models": "OST_GenericModel",
    "Lines": "OST_Lines",
}

ANNOTATION_CATEGORY_NAMES = {
    "Pipe Tags": "OST_PipeTags",
    "Duct Tags": "OST_DuctTags",
    "Pipe Accessory Tags": "OST_PipeAccessoryTags",
    "Duct Accessory Tags": "OST_DuctAccessoryTags",
    "Mechanical Equipment Tags": "OST_MechanicalEquipmentTags",
    "Plumbing Fixture Tags": "OST_PlumbingFixtureTags",
    "Generic Model Tags": "OST_GenericModelTags",
    "Structural Framing Tags": "OST_StructuralFramingTags",
    "Fabrication Pipework Tags": "OST_FabricationPipeworkTags",
    "Fabrication Ductwork Tags": "OST_FabricationDuctworkTags",
    "Fabrication Hanger Tags": "OST_FabricationHangerTags",
    "Multi-Category Tags": "OST_MultiCategoryTags",
    "Generic Annotations": "OST_GenericAnnotation",
    "Detail Items": "OST_DetailComponents",
    "Text Notes": "OST_TextNotes",
    "Dimensions": "OST_Dimensions",
}


def resolve(names):
    out = {}
    for label, enum_name in names.items():
        val = bic(enum_name)
        if val is not None:
            out[label] = val
    return out


model_categories = resolve(MODEL_CATEGORY_NAMES)
annotation_categories = resolve(ANNOTATION_CATEGORY_NAMES)
all_categories = {}
all_categories.update(model_categories)
all_categories.update(annotation_categories)


def rnd(v):
    return round(v, PRECISION)


def get_center(el):
    bbox = el.get_BoundingBox(None)
    if bbox is None:
        bbox = el.get_BoundingBox(curview)
    if bbox is None:
        return None
    c = (bbox.Max + bbox.Min) / 2
    return (rnd(c.X), rnd(c.Y), rnd(c.Z))


def is_nested_family(el):
    if isinstance(el, FamilyInstance):
        return el.SuperComponent is not None
    return False


def get_tagged_ids(tag):
    try:
        return tuple(sorted(eid(i) for i in tag.GetTaggedLocalElementIds()))
    except:
        pass
    try:
        return tuple(sorted(eid(e.Id) for e in tag.GetTaggedLocalElements()))
    except:
        pass
    try:
        return (eid(tag.TaggedLocalElementId),)
    except:
        return ()


def get_key(el, cat_label):
    if isinstance(el, IndependentTag):
        head = el.TagHeadPosition
        return (cat_label, rnd(head.X), rnd(head.Y), rnd(head.Z), get_tagged_ids(el))

    if isinstance(el, TextNote):
        c = el.Coord
        return (cat_label, rnd(c.X), rnd(c.Y), rnd(c.Z), el.Text)

    if isinstance(el, Dimension):
        center = get_center(el)
        if center is None:
            return None
        try:
            val = rnd(el.Value) if el.Value is not None else None
        except:
            val = None
        return (cat_label, center, val)

    center = get_center(el)
    if center is None:
        return None
    return (cat_label, center)


def checkout_status(el):
    if not doc.IsWorkshared:
        return CheckoutStatus.OwnedByCurrentUser
    return WorksharingUtils.GetCheckoutStatus(doc, el.Id)


def reason(el):
    """Why an element cannot be edited, from Revit's own worksharing data."""
    parts = []
    try:
        info = WorksharingUtils.GetWorksharingTooltipInfo(doc, el.Id)
        parts.append("creator={} | owner={} | lastChangedBy={}".format(
            info.Creator, info.Owner, info.LastChangedBy))
    except:
        pass
    try:
        parts.append("central status={}".format(WorksharingUtils.GetModelUpdatesStatus(doc, el.Id)))
    except:
        pass
    try:
        ws = doc.GetWorksetTable().GetWorkset(el.WorksetId)
        parts.append("workset='{}' | workset owner={} | editable={}".format(
            ws.Name, ws.Owner, ws.IsEditable))
    except:
        pass
    return "; ".join(parts) if parts else "no worksharing info available"


try:
    # ---- Select categories ----
    groups = {
        "Model": sorted(model_categories.keys()),
        "Annotation": sorted(annotation_categories.keys()),
    }

    selected = forms.SelectFromList.show(
        groups,
        group_selector_title="Type:",
        multiselect=True,
        button_name="Select Category",
        title="Overkill")

    if not selected:
        sys.exit()

    # ---- Pass 1: find duplicate groups (no edits) ----
    # cat label -> list of (keeper, [elements to delete])
    dup_sets = {}
    for cat_label in selected:
        elements = (FilteredElementCollector(doc, curview.Id)
                    .OfCategory(all_categories[cat_label])
                    .WhereElementIsNotElementType()
                    .ToElements())

        by_key = {}
        for el in elements:
            if is_nested_family(el):
                continue
            key = get_key(el, cat_label)
            if key is None:
                continue
            by_key.setdefault(key, []).append(el)

        for grp in by_key.values():
            if len(grp) < 2:
                continue
            grp = sorted(grp, key=lambda e: eid(e.Id))
            # Keep an element owned by another user if there is one (cannot delete it anyway),
            # otherwise keep the lowest ElementId so the newest copy is removed.
            keeper = grp[0]
            if doc.IsWorkshared:
                for e in grp:
                    if checkout_status(e) == CheckoutStatus.OwnedByOtherUser:
                        keeper = e
                        break
            to_del = [e for e in grp if eid(e.Id) != eid(keeper.Id)]
            dup_sets.setdefault(cat_label, []).append((keeper, to_del))

    delete_elems = []
    for cat_label, sets in dup_sets.items():
        for keeper, to_del in sets:
            delete_elems.extend(to_del)

    if not delete_elems:
        forms.show_balloon("Duplicates", "No Duplicates Found")
        sys.exit()

    # ---- Pass 2: classify ownership (read-only) ----
    editable = []       # already owned by current user, or not workshared
    need_checkout = []  # not owned by anyone yet
    blocked = []        # (element, reason)
    for el in delete_elems:
        st = checkout_status(el)
        if st == CheckoutStatus.OwnedByCurrentUser:
            editable.append(el)
        elif st == CheckoutStatus.OwnedByOtherUser:
            blocked.append((el, reason(el)))
        else:
            need_checkout.append(el)

    summary = "\n".join("{}: {}".format(k, sum(len(d) for _, d in v))
                        for k, v in sorted(dup_sets.items()))
    msg = "Delete {} duplicate(s)?\n\n{}\n\nReady to delete: {}\nNeeds checkout: {}\nBlocked (owned by others): {}".format(
        len(delete_elems), summary, len(editable), len(need_checkout), len(blocked))
    if not forms.alert(msg, yes=True, no=True):
        sys.exit()

    # ---- Pass 3: check out only what is not owned yet ----
    checkout_error = None
    if need_checkout:
        try:
            failed = WorksharingUtils.CheckoutElements(
                doc, List[ElementId]([e.Id for e in need_checkout]))
            failed_set = set(eid(i) for i in failed)
        except Exception as ex:
            checkout_error = str(ex)
            failed_set = set(eid(e.Id) for e in need_checkout)
        for el in need_checkout:
            if eid(el.Id) in failed_set:
                blocked.append((el, reason(el)))
            else:
                editable.append(el)

    # ---- Pass 4: delete in a single transaction ----
    removed_count = 0
    errors = []
    t = Transaction(doc, "Delete Duplicate Elements")
    t.Start()
    for el in editable:
        try:
            removed = doc.Delete(el.Id)
            removed_count += removed.Count
        except Exception as ex:
            errors.append("{}: {}".format(eid(el.Id), ex))
    status = t.Commit()

    # ---- Report ----
    # print("Duplicates found: {}".format(len(delete_elems)))
    # print("Delete attempts: {} | Elements removed: {}".format(len(editable), removed_count))
    # print("Blocked: {}".format(len(blocked)))
    # print("Transaction status: {}".format(status))
    if checkout_error:
        print("Checkout error: " + checkout_error)
    for el, why in blocked[:15]:
        print("BLOCKED {}: {}".format(eid(el.Id), why))
    for e in errors[:15]:
        print("Delete error: " + e)

    if status != TransactionStatus.Committed:
        forms.alert("Transaction did not commit: {}".format(status))

except SystemExit:
    pass
except Exception:
    print("SCRIPT FAILED:")
    print(traceback.format_exc())