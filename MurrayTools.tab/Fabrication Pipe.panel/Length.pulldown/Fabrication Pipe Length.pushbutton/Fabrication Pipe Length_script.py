__title__ = 'Fabrication\nPipe Length'
__doc__ = """Calculates Combined Length of Selected Fabrication Pipes."""

import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script
from pyrevit import revit
import sys

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
#path=os.path.dirname(os.path.abspath(__file__))


#Element
#PointOnElement
#Edge
#Face
#LinkedElement

#this is start of select element(s)
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
try:
    pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter("MEP Fabrication Pipework"), "Select Fabrication Pipes to collect length from")            
    Fpipework = [doc.GetElement( elId ) for elId in pipesel]
    #this is end of select element(s)

    if len(Fpipework) > 0:

        # Iterate over fabrication pipes and collect length data
        Total_Length = 0.0

        for pipe in Fpipework:
            if pipe.IsAStraight:
                len_param = pipe.Parameter[BuiltInParameter.FABRICATION_PART_LENGTH]
                if len_param:
                    Total_Length = Total_Length + len_param.AsDouble()
        # now that results are collected, print the total
        print("Linear feet of selected pipe(s) is: {}".format(Total_Length))
    else:
        forms.alert('At least one fabrication pipe must be selected.')
except:
    sys.exit()