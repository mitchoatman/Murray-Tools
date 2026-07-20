import Autodesk
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, Family, TransactionGroup, \
                              BuiltInCategory, FamilySymbol, BuiltInParameter, Reference, IndependentTag, \
                              TagMode, TagOrientation, ViewType
from Parameters.Add_SharedParameters import Shared_Params
from Autodesk.Revit.UI import TaskDialog
import os
import sys

Shared_Params()

path, filename = os.path.split(__file__)
NewFilename = 'Fabrication Hanger - FP_Rod Size.rfa'

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

FamilyName = 'Fabrication Hanger - FP_Rod Size'
FamilyType = 'RodSize'


def error_and_exit(msg):
    TaskDialog.Show("Error", msg)
    sys.exit()


def get_unique_view_name(doc, base_name):
    existing_names = set(
        v.Name for v in FilteredElementCollector(doc).OfClass(DB.View).ToElements()
        if v and hasattr(v, "Name")
    )

    if base_name not in existing_names:
        return base_name

    i = 1
    while True:
        test_name = "{}{}".format(base_name, i)
        if test_name not in existing_names:
            return test_name
        i += 1


def set_parameter_by_name(element, parameterName, value):
    param = element.LookupParameter(parameterName)
    if param and not param.IsReadOnly:
        param.Set(value)


def get_tagged_element_ids(tag):
    tagged_ids = set()

    try:
        for eid in tag.GetTaggedLocalElementIds():
            if eid and eid.IntegerValue != -1:
                tagged_ids.add(eid.IntegerValue)
        return tagged_ids
    except Exception:
        pass

    try:
        eid = tag.TaggedLocalElementId
        if eid and eid.IntegerValue != -1:
            tagged_ids.add(eid.IntegerValue)
    except Exception:
        pass

    return tagged_ids


# Handle 3D view locking with validation
if curview.ViewType == ViewType.ThreeD:
    v3d = curview

    if v3d.IsPerspective:
        error_and_exit("Tagging is not supported in perspective 3D views.")

    if "{" in v3d.Name or "}" in v3d.Name:
        new_name = get_unique_view_name(doc, "3D-TAGROD")

        t_rename = Transaction(doc, "Rename 3D View")
        t_rename.Start()
        try:
            v3d.Name = new_name
            t_rename.Commit()
        except Exception:
            t_rename.RollBack()
            error_and_exit("Failed to rename the 3D view.")

    if not v3d.IsLocked:
        t_lock = Transaction(doc, "Lock 3D View Orientation")
        t_lock.Start()
        try:
            v3d.SaveOrientationAndLock()
            t_lock.Commit()
            if not v3d.IsLocked:
                raise Exception("Lock verification failed")
        except Exception:
            t_lock.RollBack()
            error_and_exit("Failed to lock 3D view orientation. Ensure the view is not a template and retry.")

elif curview.ViewType not in [ViewType.FloorPlan, ViewType.AreaPlan, ViewType.Section]:
    error_and_exit("This script can only run in a Floor Plan, Section, or non-perspective 3D view.")


family_pathCC1 = os.path.join(path, NewFilename)

if not os.path.exists(family_pathCC1):
    error_and_exit("Family file not found:\n{}".format(family_pathCC1))


families = FilteredElementCollector(doc).OfClass(Family)
Fam_is_in_project = any(f.Name == FamilyName for f in families)


class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


tg = TransactionGroup(doc, "Add Rod Size Tags")
tg.Start()

t = Transaction(doc, 'Load Rod Size Tag Family')
t.Start()
try:
    if not Fam_is_in_project:
        fload_handler = FamilyLoaderOptionsHandler()
        doc.LoadFamily(family_pathCC1, fload_handler)
    t.Commit()
except Exception as ex:
    t.RollBack()
    tg.RollBack()
    error_and_exit("Failed to load family:\n{}\n\n{}".format(family_pathCC1, str(ex)))

hanger_tag_collector = FilteredElementCollector(doc, curview.Id) \
    .OfCategory(BuiltInCategory.OST_FabricationHangers) \
    .WhereElementIsNotElementType()

familyTypes = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_FabricationHangerTags) \
    .OfClass(FamilySymbol) \
    .ToElements()

tag_type = None

t = Transaction(doc, 'Tag Hanger Rods')
t.Start()

for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        tag_type = famtype
        if not famtype.IsActive:
            famtype.Activate()
            doc.Regenerate()
        break

if not tag_type:
    t.RollBack()
    tg.RollBack()
    error_and_exit("Could not find hanger tag family type: {}".format(FamilyType))

existing_tagged_hanger_ids = set()
existing_tags = FilteredElementCollector(doc, curview.Id).OfClass(IndependentTag).ToElements()

for tag in existing_tags:
    try:
        if tag.GetTypeId() == tag_type.Id:
            existing_tagged_hanger_ids.update(get_tagged_element_ids(tag))
    except Exception:
        pass

for hanger in hanger_tag_collector:
    try:
        if hanger.ServiceType == 56:
            if hanger.Id.IntegerValue in existing_tagged_hanger_ids:
                continue

            ref = Reference(hanger)
            hanger_location = hanger.Origin

            new_tag = IndependentTag.Create(
                doc,
                curview.Id,
                ref,
                False,
                TagMode.TM_ADDBY_CATEGORY,
                TagOrientation.Horizontal,
                hanger_location
            )

            if new_tag and new_tag.GetTypeId() != tag_type.Id:
                new_tag.ChangeTypeId(tag_type.Id)

            existing_tagged_hanger_ids.add(hanger.Id.IntegerValue)
    except Exception:
        pass

t.Commit()

hanger_collector = FilteredElementCollector(doc, curview.Id) \
    .OfCategory(BuiltInCategory.OST_FabricationHangers) \
    .WhereElementIsNotElementType() \
    .ToElements()

t = Transaction(doc, 'Sync Hanger Data')
t.Start()
for hanger in hanger_collector:
    try:
        for ancillary in hanger.GetPartAncillaryUsage():
            if ancillary.AncillaryWidthOrDiameter > 0:
                set_parameter_by_name(hanger, 'FP_Rod Size', ancillary.AncillaryWidthOrDiameter)
                break
    except Exception:
        pass
t.Commit()

tg.Assimilate()