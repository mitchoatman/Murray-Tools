import Autodesk
import os
import clr
import sys

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')

from System.Collections.Generic import List
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import (
    IFamilyLoadOptions,
    FamilySource,
    Transaction,
    TransactionGroup,
    FilteredElementCollector,
    Family,
    FamilySymbol,
    BuiltInCategory,
    BuiltInParameter,
    IndependentTag,
    TagOrientation,
    Reference,
    ElementMulticategoryFilter,
    LocationPoint,
    LocationCurve,
    XYZ,
    View3D,
    FamilyInstance
)

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView


def is_parent_family_instance(element):
    if not isinstance(element, FamilyInstance):
        return True
    return element.SuperComponent is None


# --------------------------------------------------
# CONFIG
# --------------------------------------------------
FAMILY_NAME = 'Multi Category Tag -TS_Point_Number'
FAMILY_TYPE = 'Multi Category Tag -TS_Point_Number'

SCRIPT_DIR = os.path.dirname(__file__)
FAMILY_FILE = 'Multi Category Tag -TS_Point_Number.rfa'
FAMILY_PATH = os.path.join(SCRIPT_DIR, FAMILY_FILE)

TARGET_CATEGORIES = [
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_StructuralStiffener,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_FabricationHangers,
    BuiltInCategory.OST_DuctAccessory,
]

# Candidate offsets from Revit's default tag head location
# Stronger vertical spreading first
CANDIDATE_OFFSETS = [
    XYZ(0.0, 0.0, 0.0),
    XYZ(0.0, 0.75, 0.0),
    XYZ(0.0, -0.75, 0.0),
    XYZ(0.0, 1.50, 0.0),
    XYZ(0.0, -1.50, 0.0),
    XYZ(0.0, 2.25, 0.0),
    XYZ(0.0, -2.25, 0.0),
    XYZ(0.0, 3.00, 0.0),
    XYZ(0.0, -3.00, 0.0),
    XYZ(0.0, 3.75, 0.0),
    XYZ(0.0, -3.75, 0.0),
    XYZ(0.75, 0.75, 0.0),
    XYZ(-0.75, 0.75, 0.0),
    XYZ(0.75, -0.75, 0.0),
    XYZ(-0.75, -0.75, 0.0),
    XYZ(1.50, 0.0, 0.0),
    XYZ(-1.50, 0.0, 0.0),
    XYZ(1.50, 1.50, 0.0),
    XYZ(-1.50, 1.50, 0.0),
    XYZ(1.50, -1.50, 0.0),
    XYZ(-1.50, -1.50, 0.0),
]

# Estimated tag label size in model units (feet)
# Tune these if needed for your tag family/view scale.
BASE_LABEL_WIDTH = 0.55
PER_CHAR_WIDTH = 0.18
LABEL_HEIGHT = 0.45
LABEL_PADDING_X = 0.20
LABEL_PADDING_Y = 0.12


# --------------------------------------------------
# FAMILY LOAD OPTIONS
# --------------------------------------------------
class FamilyLoaderOptionsHandler(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = FamilySource.Family
        overwriteParameterValues.Value = False
        return True


# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def get_tag_symbol(doc, family_name, type_name):
    symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    for sym in symbols:
        try:
            sym_name = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if sym.Family.Name == family_name and sym_name == type_name:
                return sym
        except:
            pass
    return None


def get_tag_point(element, view):
    loc = element.Location

    if isinstance(loc, LocationPoint):
        return loc.Point

    if isinstance(loc, LocationCurve):
        try:
            return loc.Curve.Evaluate(0.5, True)
        except:
            pass

    try:
        bbox = element.get_BoundingBox(view)
        if bbox:
            return XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0
            )
    except:
        pass

    return None


def get_all_tagged_element_ids(tags):
    tagged_ids = set()

    for tag in tags:
        try:
            for eid in tag.GetTaggedLocalElementIds():
                if eid and eid.IntegerValue != -1:
                    tagged_ids.add(eid.IntegerValue)
            continue
        except:
            pass

        try:
            eid = tag.TaggedLocalElementId
            if eid and eid.IntegerValue != -1:
                tagged_ids.add(eid.IntegerValue)
        except:
            pass

    return tagged_ids


def build_multicategory_filter(categories):
    cat_list = List[BuiltInCategory]()
    for cat in categories:
        cat_list.Add(cat)
    return ElementMulticategoryFilter(cat_list)


def add_xyz(a, b):
    return XYZ(a.X + b.X, a.Y + b.Y, a.Z + b.Z)


def get_tag_text(tag):
    try:
        txt = tag.TagText
        if txt:
            return txt.strip()
    except:
        pass
    return "TAG"


def estimate_label_box_from_head(head_point, tag_text):
    """
    Approximate label rectangle centered on TagHeadPosition.
    """
    char_count = max(len(tag_text), 1)
    width = BASE_LABEL_WIDTH + (char_count * PER_CHAR_WIDTH) + LABEL_PADDING_X
    height = LABEL_HEIGHT + LABEL_PADDING_Y

    half_w = width / 2.0
    half_h = height / 2.0

    return (
        head_point.X - half_w,
        head_point.Y - half_h,
        head_point.X + half_w,
        head_point.Y + half_h
    )


def boxes_overlap_2d(b1, b2):
    return not (
        b1[2] < b2[0] or
        b1[0] > b2[2] or
        b1[3] < b2[1] or
        b1[1] > b2[3]
    )


def collect_existing_tag_boxes(tags):
    boxes = []

    for tag in tags:
        try:
            head = tag.TagHeadPosition
            txt = get_tag_text(tag)
            box = estimate_label_box_from_head(head, txt)
            boxes.append(box)
        except:
            pass

    return boxes


def choose_tag_head_position(tag, occupied_boxes):
    base_head = tag.TagHeadPosition
    tag_text = get_tag_text(tag)

    fallback_head = base_head
    fallback_box = estimate_label_box_from_head(base_head, tag_text)

    for offset in CANDIDATE_OFFSETS:
        candidate = XYZ(
            base_head.X + offset.X,
            base_head.Y + offset.Y,
            base_head.Z + offset.Z
        )

        candidate_box = estimate_label_box_from_head(candidate, tag_text)

        overlap = False
        for existing_box in occupied_boxes:
            if boxes_overlap_2d(candidate_box, existing_box):
                overlap = True
                break

        if not overlap:
            return candidate, candidate_box

    return fallback_head, fallback_box


# --------------------------------------------------
# MAIN
# --------------------------------------------------
try:
    if isinstance(curview, View3D):
        TaskDialog.Show(
            "Unsupported View",
            "This tool does not run in 3D views.\n\n"
            "Please run it from a 2D view such as a plan, section, or elevation."
        )
        sys.exit()

    families = FilteredElementCollector(doc).OfClass(Family)
    family_in_project = any(f.Name == FAMILY_NAME for f in families)

    tg = TransactionGroup(doc, "Tag Visible Elements with TS Point Number")
    tg.Start()

    # ----------------------------------------------
    # Load family if needed
    # ----------------------------------------------
    t1 = Transaction(doc, "Load Tag Family")
    try:
        t1.Start()

        if not family_in_project:
            if not os.path.exists(FAMILY_PATH):
                raise Exception("Family file not found:\n{}".format(FAMILY_PATH))

            load_options = FamilyLoaderOptionsHandler()
            loaded = doc.LoadFamily(FAMILY_PATH, load_options)

            if not loaded:
                raise Exception("Revit could not load the family.")

        t1.Commit()
    except Exception as e:
        if t1.HasStarted():
            t1.RollBack()
        raise Exception("Failed to load family: {}".format(str(e)))

    # ----------------------------------------------
    # Get tag symbol
    # ----------------------------------------------
    tag_symbol = get_tag_symbol(doc, FAMILY_NAME, FAMILY_TYPE)
    if not tag_symbol:
        tg.RollBack()
        raise Exception(
            "Could not find tag type '{}' in family '{}'.".format(FAMILY_TYPE, FAMILY_NAME)
        )

    # ----------------------------------------------
    # Collect visible elements and existing tags
    # ----------------------------------------------
    try:
        multi_cat_filter = build_multicategory_filter(TARGET_CATEGORIES)

        elements_in_view = list(
            FilteredElementCollector(doc, curview.Id)
            .WherePasses(multi_cat_filter)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        def sort_key(elem):
            pt = get_tag_point(elem, curview)
            if pt:
                return (round(pt.X, 4), round(pt.Y, 4))
            return (0, 0)

        elements_in_view.sort(key=sort_key)

        existing_tags = (
            FilteredElementCollector(doc, curview.Id)
            .OfClass(IndependentTag)
            .ToElements()
        )

        already_tagged_ids = get_all_tagged_element_ids(existing_tags)
        occupied_tag_boxes = collect_existing_tag_boxes(existing_tags)

    except Exception as e:
        tg.RollBack()
        raise Exception("Failed to collect elements/tags: {}".format(str(e)))

    # ----------------------------------------------
    # Create tags
    # ----------------------------------------------
    t2 = Transaction(doc, "Tag Visible Elements")
    try:
        t2.Start()

        if not tag_symbol.IsActive:
            tag_symbol.Activate()
            doc.Regenerate()

        tagged_count = 0

        for element in elements_in_view:
            try:
                if not is_parent_family_instance(element):
                    continue

                if element.Id.IntegerValue in already_tagged_ids:
                    continue

                host_point = get_tag_point(element, curview)
                if not host_point:
                    continue

                ref = Reference(element)

                # Attached leader, preserve Revit default leader length
                new_tag = IndependentTag.Create(
                    doc,
                    tag_symbol.Id,
                    curview.Id,
                    ref,
                    True,
                    TagOrientation.Horizontal,
                    host_point
                )

                if not new_tag:
                    continue

                doc.Regenerate()

                final_head, final_box = choose_tag_head_position(new_tag, occupied_tag_boxes)

                try:
                    new_tag.TagHeadPosition = final_head
                    doc.Regenerate()
                except:
                    pass

                occupied_tag_boxes.append(final_box)
                already_tagged_ids.add(element.Id.IntegerValue)
                tagged_count += 1

            except:
                continue

        t2.Commit()
        tg.Assimilate()

        TaskDialog.Show("Tag Visible Elements", "Tagged: {}".format(tagged_count))

    except Exception as e:
        if t2.HasStarted():
            t2.RollBack()
        tg.RollBack()
        raise Exception("Failed while tagging elements: {}".format(str(e)))

except Exception as e:
    TaskDialog.Show("Error", str(e))
    raise