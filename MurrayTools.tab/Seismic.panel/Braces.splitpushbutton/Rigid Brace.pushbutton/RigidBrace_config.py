

from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, Structure, XYZ, FabricationPart, FabricationConfiguration
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
# Set desired family name and type name:
FamilyName = 'RIGID SEISMIC BRACE'
FamilyType = 'RIGID SEISMIC BRACE'
# Check if the family is in the project
Fam_is_in_project = any(f.Name == FamilyName for f in families)
#print("Family '{}' is in project: {}".format(FamilyName, is_in_project))

def get_parameter_value_by_name_AsDouble(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True


    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return true

pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers to distribute collected weight across")            
Fhangers = [doc.GetElement( elId ) for elId in pipesel]

family_pathCC = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\Families\Structural Stiffeners (Seismic)\RIGID SEISMIC BRACE.rfa'

t = Transaction(doc, 'Load Rigid Brace Family')
#Start Transaction
t.Start()
if Fam_is_in_project == False:
    fload_handler = FamilyLoaderOptionsHandler()
    family = doc.LoadFamily(family_pathCC, fload_handler)
# t.Commit()

#Family symbol name to place.
symbName = 'RIGID SEISMIC BRACE'

#create a filtered element collector set to Category OST_Mass and Class FamilySymbol
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_StructuralStiffener)
collector.OfClass(FamilySymbol)

famtypeitr = collector.GetElementIdIterator()
famtypeitr.Reset()

#Search Family Symbols in document.
for item in famtypeitr:
    famtypeID = item
    famsymb = doc.GetElement(famtypeID)

    #If the FamilySymbol is the name we are looking for, create a new instance.
    if famsymb.Family.Name == symbName:
        #uidoc.PostRequestForElementTypePlacement(famsymb)
        for hanger in Fhangers:
        
            try:
                if (hanger.GetRodInfo().RodCount) == 1:
                    ItmDims = hanger.GetDimensions()
                    for dta in ItmDims:
                        if dta.Name == 'Total Height':
                            HangerHeight = hanger.GetDimensionValue(dta)
                    BraceOffsetZ = HangerHeight + 0.01041666
                    # Get the bounding box of the hanger
                    bounding_box = hanger.get_BoundingBox(None)
                    if bounding_box is not None:
                        # Calculate the middle bottom point of the bounding box
                        middle_bottom_point = XYZ((bounding_box.Min.X + bounding_box.Max.X) / 2,
                                                  (bounding_box.Min.Y + bounding_box.Max.Y) / 2,
                                                  bounding_box.Min.Z)
                    new_insertion_point = XYZ(middle_bottom_point.X, middle_bottom_point.Y, middle_bottom_point.Z  + BraceOffsetZ)
                    # Create new instance
                    doc.Create.NewFamilyInstance(new_insertion_point, famsymb, DB.Structure.StructuralType.NonStructural)
            except:
                pass

            try:
                if (hanger.GetRodInfo().RodCount) > 1:
                    STName = hanger.GetRodInfo().RodCount
                    STName1 = hanger.GetRodInfo()
                    middle_elevation = get_parameter_value_by_name_AsDouble(hanger, 'Middle Elevation')
                    for n in range(STName):
                        rodloc = STName1.GetRodEndPosition(n)
                        combined_xyz = XYZ(rodloc.X, rodloc.Y, middle_elevation)
                        # Create new instance
                        doc.Create.NewFamilyInstance(combined_xyz, famsymb, DB.Structure.StructuralType.NonStructural)
            except:
                pass





t.Commit()