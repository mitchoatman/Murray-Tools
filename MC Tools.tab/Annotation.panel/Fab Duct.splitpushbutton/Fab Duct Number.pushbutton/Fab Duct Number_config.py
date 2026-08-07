# -*- coding: utf-8 -*-

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    BuiltInCategory,
    FamilySymbol,
    Family,
    IndependentTag,
    TagOrientation,
    XYZ
)
import os


# ============================================================
# FILE PATHS
# ============================================================
path, filename = os.path.split(__file__)
ALIGNED_TAG_FILENAME = 'Fabrication Duct - Number Tag - Aligned.rfa'
NONALIGNED_TAG_FILENAME = 'Fabrication Duct - Number Tag.rfa'

ALIGNED_TAG_PATH = os.path.join(path, ALIGNED_TAG_FILENAME)
NONALIGNED_TAG_PATH = os.path.join(path, NONALIGNED_TAG_FILENAME)


# ============================================================
# REVIT CONTEXT
# ============================================================
doc = __revit__.ActiveUIDocument.Document
active_view = doc.ActiveView


# ============================================================
# USER SETTINGS
# ============================================================
STRAIGHT_OFFSET_RATIO = 1.0 / 4.0   # 0.5 = half duct width; increase if tag needs to sit farther off duct
DEFAULT_RECT_WIDTH = 2.0            # feet, fallback if width param not found


# ============================================================
# FAMILY LOAD OPTIONS
# ============================================================
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


# ============================================================
# HELPERS
# ============================================================
def get_service_type_name(fab_part):
    try:
        config = DB.FabricationConfiguration.GetFabricationConfiguration(doc)
        if config:
            return config.GetServiceTypeName(fab_part.ServiceType).strip()
    except:
        pass
    return ""


def is_rectangular_fab_duct(part):
    service_type_name = get_service_type_name(part).lower()
    return "rectangular duct" in service_type_name


def get_bbox_center(element, view):
    bbox = element.get_BoundingBox(view)
    if not bbox:
        return None

    return XYZ(
        (bbox.Min.X + bbox.Max.X) / 2.0,
        (bbox.Min.Y + bbox.Max.Y) / 2.0,
        (bbox.Min.Z + bbox.Max.Z) / 2.0
    )


def get_xy_vector(p1, p2):
    return XYZ(p2.X - p1.X, p2.Y - p1.Y, 0)


def normalize_xy(vec):
    flat = XYZ(vec.X, vec.Y, 0)
    if flat.GetLength() == 0:
        return None
    return flat.Normalize()


def get_rect_width(duct):
    width_param = duct.LookupParameter("Main Primary Width")
    if width_param and width_param.HasValue:
        return width_param.AsDouble()
    return DEFAULT_RECT_WIDTH


def get_top_perpendicular(dir_xy):
    perp1 = XYZ(-dir_xy.Y, dir_xy.X, 0).Normalize()
    perp2 = XYZ(dir_xy.Y, -dir_xy.X, 0).Normalize()

    # Prefer upper Y; if tied, prefer farther left X
    if perp1.Y > perp2.Y:
        return perp1
    elif perp2.Y > perp1.Y:
        return perp2
    else:
        if perp1.X <= perp2.X:
            return perp1
        else:
            return perp2


def get_straight_tag_location(center_point, duct_direction_xy, duct_width):
    perp_dir = get_top_perpendicular(duct_direction_xy)
    offset_distance = duct_width * STRAIGHT_OFFSET_RATIO
    offset_vector = perp_dir.Multiply(offset_distance)
    return center_point.Add(offset_vector)


def create_tag(tag_symbol_id, view_id, element, tag_orientation, tag_location):
    return IndependentTag.Create(
        doc,
        tag_symbol_id,
        view_id,
        DB.Reference(element),
        False,
        tag_orientation,
        tag_location
    )


def get_tag_symbol(family_name):
    collector = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_FabricationDuctworkTags) \
        .OfClass(FamilySymbol)

    for famsymb in collector:
        if famsymb.Family.Name == family_name:
            if not famsymb.IsActive:
                famsymb.Activate()
            return famsymb
    return None


def get_existing_number_tagged_ids(view_id, target_family_names):
    existing_tags = FilteredElementCollector(doc, view_id) \
        .OfCategory(BuiltInCategory.OST_FabricationDuctworkTags) \
        .OfClass(IndependentTag)

    tagged_ids = set()

    for tag in existing_tags:
        try:
            tag_type = doc.GetElement(tag.GetTypeId())
            if not tag_type:
                continue

            family_name = tag_type.Family.Name
            if family_name not in target_family_names:
                continue

            tagged_elem_id = tag.TaggedLocalElementId
            if tagged_elem_id != DB.ElementId.InvalidElementId:
                tagged_ids.add(tagged_elem_id)
        except:
            pass

    return tagged_ids


def ensure_tag_families_loaded():
    required_family_names = {
        'Fabrication Duct - Number Tag - Aligned',
        'Fabrication Duct - Number Tag'
    }

    loaded_family_names = set(
        f.Name for f in FilteredElementCollector(doc).OfClass(Family)
    )

    if not required_family_names.issubset(loaded_family_names):
        fload_handler = FamilyLoaderOptionsHandler()
        doc.LoadFamily(ALIGNED_TAG_PATH, fload_handler)
        doc.LoadFamily(NONALIGNED_TAG_PATH, fload_handler)


# ============================================================
# MAIN
# ============================================================
t = Transaction(doc, 'Load and Place Duct Number Tags')
t.Start()

ensure_tag_families_loaded()

aligned_family_name = 'Fabrication Duct - Number Tag - Aligned'
nonaligned_family_name = 'Fabrication Duct - Number Tag'

tag_symbol_aligned = get_tag_symbol(aligned_family_name)
tag_symbol_nonaligned = get_tag_symbol(nonaligned_family_name)

if tag_symbol_aligned and tag_symbol_nonaligned:
    already_tagged_ids = get_existing_number_tagged_ids(
        active_view.Id,
        {aligned_family_name, nonaligned_family_name}
    )

    duct_collector = FilteredElementCollector(doc, active_view.Id) \
        .OfCategory(BuiltInCategory.OST_FabricationDuctwork) \
        .WhereElementIsNotElementType()

    tagged_straights_count = 0
    tagged_fittings_count = 0
    skipped_count = 0
    already_tagged_count = 0

    for part in duct_collector:
        try:
            if part.Id in already_tagged_ids:
                already_tagged_count += 1
                continue

            if not is_rectangular_fab_duct(part):
                skipped_count += 1
                continue

            center_point = get_bbox_center(part, active_view)
            if not center_point:
                skipped_count += 1
                continue

            # STRAIGHT RECTANGULAR DUCT
            if part.IsAStraight():
                connector_manager = part.ConnectorManager
                if not connector_manager or connector_manager.Connectors.Size < 2:
                    skipped_count += 1
                    continue

                connectors = list(connector_manager.Connectors)
                conn1_origin = connectors[0].Origin
                conn2_origin = connectors[1].Origin

                duct_direction_xy = normalize_xy(get_xy_vector(conn1_origin, conn2_origin))
                if not duct_direction_xy:
                    skipped_count += 1
                    continue

                duct_width = get_rect_width(part)
                tag_location = get_straight_tag_location(center_point, duct_direction_xy, duct_width)

                create_tag(
                    tag_symbol_aligned.Id,
                    active_view.Id,
                    part,
                    TagOrientation.Horizontal,
                    tag_location
                )

                tagged_straights_count += 1

            # NON-STRAIGHT RECTANGULAR PARTS / FITTINGS
            else:
                create_tag(
                    tag_symbol_nonaligned.Id,
                    active_view.Id,
                    part,
                    TagOrientation.Horizontal,
                    center_point
                )

                tagged_fittings_count += 1

        except Exception as e:
            print("Error tagging part {}: {}".format(part.Id, str(e)))
            skipped_count += 1
            continue

t.Commit()

# print("Duct number tagging completed!")
# print("Tagged {} straight rectangular ducts".format(tagged_straights_count))
# print("Tagged {} rectangular fittings".format(tagged_fittings_count))
# print("Skipped {} already-tagged parts".format(already_tagged_count))
# print("Skipped {} other parts".format(skipped_count))