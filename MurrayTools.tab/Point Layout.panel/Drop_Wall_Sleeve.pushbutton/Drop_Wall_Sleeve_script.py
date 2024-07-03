
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Family, BuiltInCategory, FamilySymbol, LocationCurve, Transaction
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger
import re
from math import atan2, degrees

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set desired family name and type name:
FamilyName = 'DR-WS'
FamilyType = 'DR-WS'
# Check if the family is in the project
Fam_is_in_project = any(f.Name == FamilyName for f in families)
#print("Family '{}' is in project: {}".format(FamilyName, is_in_project))

family_pathCC = r'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES\Pipe Accessories\Sleeves\DR-WS.rfa'

t = Transaction(doc, 'Load Trimble Wall Sleeve Family')
#Start Transaction
t.Start()
if Fam_is_in_project == False:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_pathCC, fload_handler)

#create a filtered element collector set to Category OST_PipeAccessory and Class FamilySymbol
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_PipeAccessory)
collector.OfClass(FamilySymbol)

famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()

#Search Family Symbols in document.
for item in famtypeitr:
    famtypeID = item
    famsymb = doc.GetElement(famtypeID)

if famsymb:
    famsymb.Activate()
    doc.Regenerate()

t.Commit()

#Family symbol name to place.
symbName = 'DR-WS'


def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def get_parameter_value_by_name_AsDouble(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

# Function to prompt user to select an MEP Fabrication Pipe
def select_fabrication_pipe():
    selection = uidoc.Selection
    pipe_ref = selection.PickObject(ObjectType.Element, "Select an MEP Fabrication Pipe")
    pipe = doc.GetElement(pipe_ref.ElementId)
    return pipe

# Function to prompt user to pick a point
def pick_point():
    picked_point = uidoc.Selection.PickPoint("Pick a point along the centerline of the pipe")
    return picked_point

# Function to get the centerline curve of the pipe
def get_pipe_centerline(pipe):
    pipe_location = pipe.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    else:
        raise Exception("The selected element does not have a valid centerline.")

# Function to project point onto the centerline curve
def project_point_on_curve(point, curve):
    result = curve.Project(point)
    return result.XYZPoint

# Main function
def place_and_modify_family(pipe, famsymb):
    # Get the centerline of the pipe
    centerline_curve = get_pipe_centerline(pipe)
    
    # Prompt user to pick a point along the centerline
    picked_point = pick_point()
    
    # Project the picked point onto the centerline curve
    projected_point = project_point_on_curve(picked_point, centerline_curve)
    
    # Create new family in model at picked point
    new_family_instance = doc.Create.NewFamilyInstance(projected_point, famsymb, DB.Structure.StructuralType.NonStructural)
    
    # Set family diameter based on selected pipe
    diameter = float(re.sub(r'[^\d.]', '', get_parameter_value_by_name_AsString(pipe, 'Overall Size'))) / 12 + 0.0833333
    set_parameter_by_name(new_family_instance, "Diameter", diameter)
    
    # Get location of inserted family
    insertion_point = new_family_instance.Location.Point
    
    # Get connector location and angle on pipe
    pipe_connector = pipe.ConnectorManager.Connectors
    connector1, connector2 = list(pipe_connector)
    vec_x = connector2.Origin.X - connector1.Origin.X
    vec_y = connector2.Origin.Y - connector1.Origin.Y
    angle = atan2(vec_y, vec_x)
    axis = DB.Line.CreateBound(insertion_point, DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1))
    
    # Set rotation on new family placed in model
    DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)
    
    # Set FP parameters on new family placed in model
    set_parameter_by_name(new_family_instance, 'FP_Service Name', get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Name'))
    set_parameter_by_name(new_family_instance, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Abbreviation'))


# Enter while loop for family placement
while True:
    try:
        # Start transaction for each family placement
        t = Transaction(doc, 'Place Trimble Wall Sleeve Family')
        t.Start()
        
        # Select the MEP Fabrication Pipe
        pipe = select_fabrication_pipe()
        place_and_modify_family(pipe, famsymb)
        
        t.Commit()
        
    except Exception as e:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        break
