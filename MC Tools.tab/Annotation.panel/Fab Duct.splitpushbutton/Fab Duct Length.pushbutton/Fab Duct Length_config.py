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
ALIGNED_TAG_FILENAME = 'Fabrication Duct - Duct Length Tag - Aligned.rfa'
NONALIGNED_TAG_FILENAME = 'Fabrication Duct - Duct Length Tag.rfa'

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
ROUND_OFFSET_RATIO = 1.0 / 4.0
RECT_OFFSET_RATIO = 1.0 / 4.0
DEFAULT_RECT_WIDTH = 2.0
DIRECTION_DOMINANCE_RATIO = 1.25


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


def get_connector_size(connector):
    try:
        shape = connector.Shape
        if shape == DB.ConnectorProfileType.Round:
            return (connector.Radius, connector.Radius)
        elif shape == DB.ConnectorProfileType.Rectangular or shape == DB.ConnectorProfileType.Oval:
            return (connector.Width, connector.Height)
    except:
        pass
    return None


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


def classify_direction(dir_xy):
    abs_x = abs(dir_xy.X)
    abs_y = abs(dir_xy.Y)

    if abs_x == 0 and abs_y == 0:
        return "none"

    if abs_x >= abs_y * DIRECTION_DOMINANCE_RATIO:
        return "horizontal"
    elif abs_y >= abs_x * DIRECTION_DOMINANCE_RATIO:
        return "vertical"
    else:
        return "diagonal"


def get_rect_width(duct):
    width_param = duct.LookupParameter("Main Primary Width")
    if width_param and width_param.HasValue:
        return width_param.AsDouble()
    return DEFAULT_RECT_WIDTH


def get_rect_offset_direction(dir_xy, direction_class):
    if direction_class == "horizontal":
        return XYZ(0, -1, 0)   # down

    if direction_class == "vertical":
        return XYZ(1, 0, 0)    # right

    # diagonal: choose lower; if tied, choose left
    perp1 = XYZ(-dir_xy.Y, dir_xy.X, 0).Normalize()
    perp2 = XYZ(dir_xy.Y, -dir_xy.X, 0).Normalize()

    if perp1.Y < perp2.Y:
        return perp1
    elif perp2.Y < perp1.Y:
        return perp2
    else:
        return perp1 if perp1.X <= perp2.X else perp2


def get_preferred_round_end(conn1_origin, conn2_origin, direction_class):
    # Horizontal: left end
    if direction_class == "horizontal":
        return conn1_origin if conn1_origin.X <= conn2_origin.X else conn2_origin

    # Vertical: bottom end
    if direction_class == "vertical":
        return conn1_origin if conn1_origin.Y <= conn2_origin.Y else conn2_origin

    # Diagonal: left end
    return conn1_origin if conn1_origin.X <= conn2_origin.X else conn2_origin


def get_round_tag_location(center_point, conn1_origin, conn2_origin, duct_length, direction_class):
    preferred_end = get_preferred_round_end(conn1_origin, conn2_origin, direction_class)

    direction_to_end = get_xy_vector(center_point, preferred_end)
    direction_xy = normalize_xy(direction_to_end)
    if not direction_xy:
        return center_point

    offset_distance = duct_length * ROUND_OFFSET_RATIO
    offset_vector = direction_xy.Multiply(offset_distance)
    return center_point.Add(offset_vector)


def get_rect_tag_location(center_point, duct_direction_xy, duct_width, direction_class):
    offset_dir = get_rect_offset_direction(duct_direction_xy, direction_class)
    offset_distance = duct_width * RECT_OFFSET_RATIO
    offset_vector = offset_dir.Multiply(offset_distance)
    return center_point.Add(offset_vector)


def get_transition_tag_orientation(direction_xy):
    if abs(direction_xy.Y) > abs(direction_xy.X):
        return TagOrientation.Vertical
    return TagOrientation.Horizontal


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


def get_existing_tagged_ids(view_id, target_family_names):
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
    family_name = 'Fabrication Duct - Duct Length Tag'
    families = FilteredElementCollector(doc).OfClass(Family)
    is_loaded = any(f.Name == family_name for f in families)

    if not is_loaded:
        fload_handler = FamilyLoaderOptionsHandler()
        doc.LoadFamily(ALIGNED_TAG_PATH, fload_handler)
        doc.LoadFamily(NONALIGNED_TAG_PATH, fload_handler)


# ============================================================
# MAIN
# ============================================================
t = Transaction(doc, 'Load and Place Duct Length Tags')
t.Start()

ensure_tag_families_loaded()

tag_symbol_aligned = get_tag_symbol('Fabrication Duct - Duct Length Tag - Aligned')
tag_symbol_nonaligned = get_tag_symbol('Fabrication Duct - Duct Length Tag')

if tag_symbol_aligned and tag_symbol_nonaligned:
    already_tagged_duct_ids = get_existing_tagged_ids(
    active_view.Id,
    {
        'Fabrication Duct - Duct Length Tag - Aligned',
        'Fabrication Duct - Duct Length Tag'
    }
)

    duct_collector = FilteredElementCollector(doc, active_view.Id) \
        .OfCategory(BuiltInCategory.OST_FabricationDuctwork) \
        .WhereElementIsNotElementType()

    tagged_straights_count = 0
    tagged_transitions_count = 0
    skipped_count = 0
    already_tagged_count = 0

    for duct in duct_collector:
        try:
            if duct.Id in already_tagged_duct_ids:
                already_tagged_count += 1
                continue

            connector_manager = duct.ConnectorManager
            if not connector_manager or connector_manager.Connectors.Size < 2:
                skipped_count += 1
                continue

            connectors = list(connector_manager.Connectors)
            conn1 = connectors[0]
            conn2 = connectors[1]
            conn1_origin = conn1.Origin
            conn2_origin = conn2.Origin

            center_point = get_bbox_center(duct, active_view)
            if not center_point:
                skipped_count += 1
                continue

            duct_direction_xy = normalize_xy(get_xy_vector(conn1_origin, conn2_origin))
            if not duct_direction_xy:
                skipped_count += 1
                continue

            direction_class = classify_direction(duct_direction_xy)
            duct_length = conn1_origin.DistanceTo(conn2_origin)

            if duct.IsAStraight():
                service_type_name = get_service_type_name(duct).lower()

                if service_type_name == "round duct" or "round duct" in service_type_name:
                    tag_location = get_round_tag_location(
                        center_point,
                        conn1_origin,
                        conn2_origin,
                        duct_length,
                        direction_class
                    )
                else:
                    duct_width = get_rect_width(duct)
                    tag_location = get_rect_tag_location(
                        center_point,
                        duct_direction_xy,
                        duct_width,
                        direction_class
                    )

                create_tag(
                    tag_symbol_aligned.Id,
                    active_view.Id,
                    duct,
                    TagOrientation.Horizontal,
                    tag_location
                )

                tagged_straights_count += 1

            else:
                if connector_manager.Connectors.Size != 2:
                    skipped_count += 1
                    continue

                size1 = get_connector_size(conn1)
                size2 = get_connector_size(conn2)

                if size1 is None or size2 is None:
                    skipped_count += 1
                    continue

                if size1 == size2:
                    skipped_count += 1
                    continue

                tag_orientation = get_transition_tag_orientation(duct_direction_xy)

                create_tag(
                    tag_symbol_nonaligned.Id,
                    active_view.Id,
                    duct,
                    tag_orientation,
                    center_point
                )

                tagged_transitions_count += 1

        except Exception as e:
            print("Error tagging duct {}: {}".format(duct.Id, str(e)))
            skipped_count += 1
            continue

t.Commit()

# print("Duct tagging completed!")
# print("Tagged {} new straight ducts".format(tagged_straights_count))
# print("Tagged {} new transitions/offsets".format(tagged_transitions_count))
# print("Skipped {} already-tagged ducts".format(already_tagged_count))
# print("Skipped {} other fittings".format(skipped_count))