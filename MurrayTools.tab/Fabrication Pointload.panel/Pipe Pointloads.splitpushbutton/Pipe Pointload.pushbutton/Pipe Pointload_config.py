__title__ = 'Pipe\nPointload'
__doc__ = """Calculates the combined weight of selected Fabrication Pipes
and divides that weight across the selected Fabrication Hangers.
1. Run Script.
2. Select Fabrication Pipes you wish to collect weight from.
3. Select Fabrication Hangers you wish to distribute the collected weight across.

Planned Improvements:
Add Functionality for Trapeze Hangers
Add Tagging functions
"""
__highlight__ = 'new'


import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import re
from SharedParam.Add_Parameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

fraction_pattern = re.compile(r"^(?P<num>[0-9]+)/(?P<den>[0-9]+)$")


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
CustomISelectionFilter("MEP Fabrication Pipework"), "Select Fabrication Pipes to collect weight from")            
Fpipework = [doc.GetElement( elId ) for elId in pipesel]

    

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


#start of defining functions to use
def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def get_parameter_numvalue_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()
    
def round_up(n, decimals=0):
    multiplier = 10 ** decimals
    return math.ceil(n * multiplier) / multiplier
#end of defining functions to use

if len(Fpipework) > 0:

    # Iterate over fabrication pipes and collect length data
    Total_Weight = 0.0
    
    for pipe in Fpipework:
        if pipe.ItemCustomId == (2041):
            piperad = (get_parameter_numvalue_by_name(pipe, 'Main Primary Diameter') * 6)
            pipelength = get_parameter_numvalue_by_name(pipe, 'Length')
            pweight_param = get_parameter_value_by_name(pipe, 'Weight')
            pipelb_param = float (pweight_param.replace(" lbm", ""))
            B = (piperad * piperad * 3.14159)
            C = ((pipelength * 12) * B)
            D = (C / 231)
            Z = (D * 8.34)
            F = (pipelb_param + Z)
            #print(piperad)
            #print(B)
            #print(C)
            #print(D)
            #print(Z)
            Total_Weight = Total_Weight + F
        else:
            fweight_param = get_parameter_value_by_name(pipe, 'Weight')
            fittinglb_param = float (fweight_param.replace(" lbm", ""))
            Total_Weight = Total_Weight + fittinglb_param

    if len(Fhangers) > 0:
        Hanger_Count = 0.0
        
        for hanger in Fhangers:
            hangercount = 1
            Hanger_Count = Hanger_Count + hangercount
    else:
        forms.alert('At least one fabrication hanger must be selected.')
        
    pointload = ((Total_Weight / Hanger_Count) / 10)

    t = Transaction(doc, 'Write Pointload Info')
    #Start Transaction
    t.Start()

    for whanger in Fhangers:
        numofrods = whanger.GetRodInfo().RodCount
        if numofrods > 0:
            roundedpointload = round_up(pointload) / numofrods
            set_parameter_by_name(whanger,"FP_Pointload", roundedpointload)
        
    #End Transaction
    t.Commit()

else:
    from pyrevit import forms
    forms.alert('At least one fabrication Pipe or Hanger was not selected.')