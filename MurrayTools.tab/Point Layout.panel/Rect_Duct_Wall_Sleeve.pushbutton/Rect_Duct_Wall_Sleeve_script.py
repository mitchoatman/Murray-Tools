from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Family, BuiltInCategory, FamilySymbol, LocationCurve, Transaction
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsDouble
import re
from math import atan2, degrees
from fractions import Fraction
import os

doc = __revit__.ActiveUIDocument.Document

selection = [doc.GetElement(id) for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]

if selection:
    t = Transaction(doc, "Update FP Parameters")
    t.Start()

    for x in selection:
        try:
            oldheight = get_parameter_value_by_name_AsDouble(x, 'Height')
            oldwidth = get_parameter_value_by_name_AsDouble(x, 'Width')
            # Set FP parameters on new family placed in model
            set_parameter_by_name(x, 'Height', oldwidth)
            set_parameter_by_name(x, 'Width', oldheight)
        except:
            pass
    t.Commit()


else:
    path, filename = os.path.split(__file__)
    NewFilename = '\RWS.rfa'

    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    active_view = doc.ActiveView

    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_Duct-Wall-Sleeve.txt')

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    if not os.path.exists(filepath):
        with open(filepath, 'w') as f:
            f.write('1')

    with open(filepath, 'r') as f:
        AnnularSpace = float(f.read())

    # Get the associated level of the active view
    level = active_view.GenLevel

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
    FamilyName = 'RWS'
    FamilyType = 'RWS'
    # Check if the family is in the project
    Fam_is_in_project = any(f.Name == FamilyName for f in families)

    family_pathCC = path + NewFilename

    t = Transaction(doc, 'Load Trimble Wall Sleeve Family')
    t.Start()
    if not Fam_is_in_project:
        fload_handler = FamilyLoaderOptionsHandler()
        family = doc.LoadFamily(family_pathCC, fload_handler)

    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctAccessory).OfClass(FamilySymbol)

    # Filter family symbols by family name and type name
    famsymb = None
    for fs in collector:
        if fs.Family.Name == FamilyName and fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType:
            famsymb = fs
            break

    if famsymb:
        famsymb.Activate()
        doc.Regenerate()

    t.Commit()

    symbName = 'Rectangular Sleeve'

    def select_fabrication_duct():
        selection = uidoc.Selection
        pipe_ref = selection.PickObject(ObjectType.Element, "Select a round MEP Fabrication Duct")
        duct = doc.GetElement(pipe_ref.ElementId)
        return duct

    def pick_point():
        picked_point = uidoc.Selection.PickPoint("Pick a point along the centerline of the Duct")
        return picked_point

    def get_duct_centerline(duct):
        pipe_location = duct.Location
        if isinstance(pipe_location, LocationCurve):
            return pipe_location.Curve
        else:
            raise Exception("The selected element does not have a valid centerline.")

    def project_point_on_curve(point, curve):
        result = curve.Project(point)
        return result.XYZPoint

    def place_and_modify_family(duct, famsymb):
        centerline_curve = get_duct_centerline(duct)
        picked_point = pick_point()
        
        # Project the picked point onto the duct centerline to get the Z coordinate
        projected_point = project_point_on_curve(picked_point, centerline_curve)
        
        # Create the insertion point using the X and Y from the picked point and Z from the projected point
        insertion_point = DB.XYZ(picked_point.X, picked_point.Y, projected_point.Z)
        
        new_family_instance = doc.Create.NewFamilyInstance(insertion_point, famsymb, DB.Structure.StructuralType.NonStructural)

        def extract_width_height(duct):
            # Get the parameter value and remove the 'x' and '"'
            overall_size = get_parameter_value_by_name_AsString(duct, 'Overall Size')
            
            # Use regular expression to extract two numbers separated by 'x'
            match = re.match(r'(\d+)"x(\d+)"', overall_size)
            
            if match:
                # Extract width and height as numbers
                width = float(match.group(1))
                height = float(match.group(2))
                return width, height
            else:
                raise ValueError("Overall Size format is invalid")
        width, height = extract_width_height(duct)
        set_parameter_by_name(new_family_instance, 'Width', width  /12 + (AnnularSpace / 12))
        set_parameter_by_name(new_family_instance, 'Height', height  /12 + (AnnularSpace / 12))  

        # Get connector locations
        duct_connectors = list(duct.ConnectorManager.Connectors)
        connector1, connector2 = duct_connectors[0], duct_connectors[1]

        # Calculate distances to the picked_point
        distance1 = picked_point.DistanceTo(connector1.Origin)
        distance2 = picked_point.DistanceTo(connector2.Origin)

        # Determine the nearest connector
        if distance1 < distance2:
            connector1, connector2 = connector1, connector2
        else:
            connector1, connector2 = connector2, connector1

        # Calculate vector components and angle
        vec_x = connector2.Origin.X - connector1.Origin.X
        vec_y = connector2.Origin.Y - connector1.Origin.Y
        angle = atan2(vec_y, vec_x)
        axis = DB.Line.CreateBound(insertion_point, DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1))
        
        # Set rotation on new family placed in model
        DB.ElementTransformUtils.RotateElement(doc, new_family_instance.Id, axis, angle)
        
        # Set FP parameters on new family placed in model
        set_parameter_by_name(new_family_instance, 'FP_Service Name', get_parameter_value_by_name_AsString(duct, 'Fabrication Service Name'))
        schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
        schedule_level_param.Set(level.Id)


    while True:
        try:
            t = Transaction(doc, 'Place Trimble Wall Sleeve Family')
            t.Start()
            
            duct = select_fabrication_duct()
            place_and_modify_family(duct, famsymb)
            
            t.Commit()
            
        except Exception as e:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            break
