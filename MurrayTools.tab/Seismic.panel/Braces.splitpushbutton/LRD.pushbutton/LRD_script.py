from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, Structure, XYZ, LocationCurve, TransactionGroup
from Autodesk.Revit.UI.Selection import ObjectType
import os
from math import atan2, pi

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

path, filename = os.path.split(__file__)
NewFilename = r'\LRD.rfa'

# Get the current view's associated level
curview = doc.ActiveView
level = curview.GenLevel

# Family setup
FamilyName = 'LRD'

# Available LRD types mapped to pipe sizes (in inches)
LRD_TYPES = {
    1: 'LRD1',
    1.25: 'LRD1.25',
    1.5: 'LRD1.5',
    2: 'LRD2',
    2.5: 'LRD2.5',
    3: 'LRD3',
    4: 'LRD4',
    5: 'LRD5',
    6: 'LRD6',
    8: 'LRD8',
    10: 'LRD10',
    12: 'LRD12',
    14: 'LRD14'
}

# Check if family is in project
families = FilteredElementCollector(doc).OfClass(Family)
target_family = next((f for f in families if f.Name == FamilyName), None)
Fam_is_in_project = target_family is not None

# Family loading options
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

def select_fabrication_pipe():
    pipe_ref = uidoc.Selection.PickObject(ObjectType.Element, 
                                        "Select an MEP Fabrication Pipe")
    return doc.GetElement(pipe_ref.ElementId)

def get_pipe_centerline(pipe):
    location = pipe.Location
    if isinstance(location, LocationCurve):
        return location.Curve
    raise Exception("Selected element has no valid centerline")

def project_point_on_curve(point, curve):
    return curve.Project(point).XYZPoint

def pick_insertion_point():
    return uidoc.Selection.PickPoint("Pick a point for LRD placement")

def get_pipe_size(pipe):
    # Get pipe size in inches (assuming 'Size' parameter exists in fabrication pipes)
    size_param = pipe.LookupParameter('Size')
    if size_param and size_param.StorageType == DB.StorageType.String:
        # Extract numeric value from size string (e.g., "1\"" -> 1.0)
        size_str = size_param.AsString()
        try:
            size = float(''.join(filter(str.isdigit, size_str.split('.')[0])))
            return size
        except:
            pass
    return None

def get_lrd_type_for_size(size):
    # Find the closest matching LRD type
    if size is None:
        return 'LRD1'  # Default to smallest size if size can't be determined
    available_sizes = sorted(LRD_TYPES.keys())
    for s in available_sizes:
        if size <= s:
            return LRD_TYPES[s]
    return LRD_TYPES[available_sizes[-1]]  # Return largest size if pipe is bigger than max

try:
    # Select pipe
    pipe = select_fabrication_pipe()
    pipe_size = get_pipe_size(pipe)
    FamilyType = get_lrd_type_for_size(pipe_size)
    
    # Get pipe centerline
    centerline = get_pipe_centerline(pipe)
    
    # Start transaction group
    with TransactionGroup(doc, "Place LRD") as tg:
        tg.Start()

        # Load family if not present
        if not Fam_is_in_project:
            t = Transaction(doc, 'Load LRD Family')
            t.Start()
            fload_handler = FamilyLoaderOptionsHandler()
            family_path = path + NewFilename
            target_family = doc.LoadFamily(family_path, fload_handler)
            t.Commit()

        # Get family symbol for the specific type
        family_symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)
        target_famtype = next((fs for fs in family_symbols 
                             if fs.Family.Name == FamilyName and 
                             fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType), None)

        if target_famtype:
            t = Transaction(doc, 'Place LRD')
            t.Start()
            
            # Activate symbol
            target_famtype.Activate()
            doc.Regenerate()

            # Get user point and project to pipe Z
            picked_point = pick_insertion_point()
            projected_point = project_point_on_curve(picked_point, centerline)
            insertion_point = DB.XYZ(picked_point.X, picked_point.Y, projected_point.Z)

            # Place family
            new_lrd = doc.Create.NewFamilyInstance(
                insertion_point,
                target_famtype,
                DB.Structure.StructuralType.NonStructural
            )

            # Rotate to match pipe direction with 90-degree adjustment
            pipe_connectors = list(pipe.ConnectorManager.Connectors)
            vec_x = pipe_connectors[1].Origin.X - pipe_connectors[0].Origin.X
            vec_y = pipe_connectors[1].Origin.Y - pipe_connectors[0].Origin.Y
            angle = atan2(vec_y, vec_x) + (pi / 2)  # Add 90 degrees (pi/2 radians)
            axis = DB.Line.CreateBound(insertion_point, 
                                     DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1))
            DB.ElementTransformUtils.RotateElement(doc, new_lrd.Id, axis, angle)

            # Set Schedule Level parameter
            schedule_level_param = new_lrd.LookupParameter("Schedule Level")
            if schedule_level_param:
                schedule_level_param.Set(level.Id)
            else:
                print "Warning: Schedule Level parameter not found in LRD family"

            t.Commit()
        else:
            print "Could not find LRD type: %s" % FamilyType
        tg.Assimilate()

except Exception, e:
    print "Error: %s" % str(e)