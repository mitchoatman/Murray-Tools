import Autodesk
import os
import clr

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
    View3D
)
from Autodesk.Revit.DB import FamilyInstance
import sys

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
        # Revit versions that support multiple tagged local ids
        try:
            for eid in tag.GetTaggedLocalElementIds():
                if eid and eid.IntegerValue != -1:
                    tagged_ids.add(eid.IntegerValue)
            continue
        except:
            pass

        # Fallback for older/single-tag behavior
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


# --------------------------------------------------
# MAIN
# --------------------------------------------------
try:
    # ----------------------------------------------
    # Block all 3D views
    # ----------------------------------------------
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
    # Collect visible elements in active view
    # ----------------------------------------------
    try:
        multi_cat_filter = build_multicategory_filter(TARGET_CATEGORIES)

        elements_in_view = (
            FilteredElementCollector(doc, curview.Id)
            .WherePasses(multi_cat_filter)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        existing_tags = (
            FilteredElementCollector(doc, curview.Id)
            .OfClass(IndependentTag)
            .ToElements()
        )

        already_tagged_ids = get_all_tagged_element_ids(existing_tags)

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

                tag_point = get_tag_point(element, curview)
                if not tag_point:
                    continue

                ref = Reference(element)

                IndependentTag.Create(
                    doc,
                    tag_symbol.Id,
                    curview.Id,
                    ref,
                    False,
                    TagOrientation.Horizontal,
                    tag_point
                )

                # Prevent duplicates in the same run
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