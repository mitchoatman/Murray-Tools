# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import (
    BuiltInCategory,
    CategoryType,
    ElementId,
    FilteredElementCollector,
    LocationCurve,
    LocationPoint,
    StorageType,
    XYZ,
)
from pyrevit import script
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# ------------------------------------------------------------
# SETTINGS
# ------------------------------------------------------------

# 10 miles in feet.
# Anything this far out is unsafe for production/NWC export.
THRESHOLD_FT = 52800.0

origin = XYZ(0, 0, 0)

# Skip non-physical / noisy categories from the broad model scan.
SKIP_CATEGORY_NAMES = set([
    "Views",
    "Cameras",
    "Room Tags",
    "Space Tags",
    "Area Tags",
    "Door Tags",
    "Window Tags",
    "Pipe Tags",
    "Mechanical Equipment Tags",
    "Generic Model Tags",
    "Text Notes",
    "Dimensions",
    "Sheets",
    "Viewports",
    "Schedules",
    "Revision Clouds",
    "Matchline",
    "Scope Boxes",
    "Levels",
    "Grids",
    "Reference Planes",
    "Project Information"
])

# Categories that should always be checked for your NWC/fabrication workflow.
TARGET_BUILTIN_CATEGORIES = set([
    BuiltInCategory.OST_FabricationPipework,
    BuiltInCategory.OST_FabricationHangers,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_PlumbingFixtures
])


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------

def safe_category_name(el):
    try:
        if el.Category:
            return el.Category.Name
    except:
        pass
    return "No Category"


def safe_element_name(el):
    try:
        return el.Name
    except:
        return ""


def safe_type_name(el):
    try:
        type_id = el.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            typ = doc.GetElement(type_id)
            if typ:
                return typ.Name
    except:
        pass
    return ""


def safe_level_name(el):
    try:
        level_id = el.LevelId
        if level_id and level_id != ElementId.InvalidElementId:
            lvl = doc.GetElement(level_id)
            if lvl:
                return lvl.Name
    except:
        pass
    return "N/A"


def get_param_value(el, param_name):
    try:
        p = el.LookupParameter(param_name)
        if not p:
            return ""
        if p.StorageType == StorageType.String:
            return p.AsString() or ""
        if p.StorageType == StorageType.Integer:
            return str(p.AsInteger())
        if p.StorageType == StorageType.Double:
            return "{:.3f}".format(p.AsDouble())
        if p.StorageType == StorageType.ElementId:
            return str(p.AsElementId().IntegerValue)
    except:
        pass
    return ""


def get_point_from_element(el):
    """
    Lightweight point extraction:
    1. LocationPoint
    2. LocationCurve start point
    3. Bounding box center fallback

    Avoids heavy geometry calls for large production models.
    """
    try:
        loc = el.Location

        if isinstance(loc, LocationPoint):
            return loc.Point

        if isinstance(loc, LocationCurve):
            crv = loc.Curve
            if crv:
                return crv.GetEndPoint(0)
    except:
        pass

    try:
        bbox = el.get_BoundingBox(None)
        if bbox:
            return XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0
            )
    except:
        pass

    return None


def should_skip_element(el):
    """
    Skip known non-model/noisy elements.
    Keep physical model categories and fabrication categories.
    """
    cat_name = safe_category_name(el)

    if cat_name in SKIP_CATEGORY_NAMES:
        return True

    try:
        cat = el.Category
        if not cat:
            return True

        # Exclude annotation/detail categories when possible.
        if cat.CategoryType != CategoryType.Model:
            return True

    except:
        return True

    return False


def distance_from_origin(pt):
    return pt.DistanceTo(origin)


# ------------------------------------------------------------
# MAIN SCAN
# ------------------------------------------------------------

bad = []
scanned_count = 0
skipped_count = 0

collector = FilteredElementCollector(doc).WhereElementIsNotElementType()

for el in collector:
    try:
        if should_skip_element(el):
            skipped_count += 1
            continue

        pt = get_point_from_element(el)
        if not pt:
            skipped_count += 1
            continue

        scanned_count += 1
        dist = distance_from_origin(pt)

        if dist > THRESHOLD_FT:
            bad.append((el, dist, pt))

    except:
        skipped_count += 1
        continue


# ------------------------------------------------------------
# OUTPUT
# ------------------------------------------------------------

output.print_md("# NWC Rogue Element Scanner")
output.print_md("**Threshold:** `{:.0f} ft`".format(THRESHOLD_FT))
output.print_md("**Model elements scanned:** `{}`".format(scanned_count))
output.print_md("**Skipped / non-model / no location:** `{}`".format(skipped_count))
output.print_md("---")

if not bad:
    output.print_md("## ✅ No physical model elements found outside threshold")
else:
    output.print_md("## 🚨 Physical Model Elements Outside Design Limits")
    output.print_md("These are the elements most likely to break NWC export.")
    output.print_md("")

    ids = List[ElementId]()

    for el, dist, pt in bad:
        ids.Add(el.Id)

        cat_name = safe_category_name(el)
        elem_name = safe_element_name(el)
        type_name = safe_type_name(el)
        level_name = safe_level_name(el)

        service = get_param_value(el, "FP_Service Name")
        if not service:
            service = get_param_value(el, "Service Name")
        if not service:
            service = get_param_value(el, "ServiceName")

        item_number = get_param_value(el, "Item Number")
        cid = get_param_value(el, "FP_CID")

        output.print_md(
            "- Element {} | **{}** | Type: `{}` | Name: `{}` | Level: `{}` | Dist: **{:.2f} ft** | X:{:.2f} Y:{:.2f} Z:{:.2f} | Service: `{}` | Item: `{}` | CID: `{}`".format(
                output.linkify(el.Id),
                cat_name,
                type_name,
                elem_name,
                level_name,
                dist,
                pt.X,
                pt.Y,
                pt.Z,
                service,
                item_number,
                cid
            )
        )

    try:
        uidoc.Selection.SetElementIds(ids)
        output.print_md("")
        output.print_md("✅ Flagged elements have been selected in Revit.")
    except:
        output.print_md("")
        output.print_md("⚠ Could not auto-select flagged elements.")