
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import re
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsValueString
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
CustomISelectionFilter("MEP Fabrication Ductwork"), "Select Fabrication Ducts to collect weight from")            
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

def round_up(n, decimals=0):
    multiplier = 10 ** decimals
    return math.ceil(n * multiplier) / multiplier
#end of defining functions to use

if len(Fpipework) > 0:

    # Iterate over fabrication pipes and collect length data
    Total_Weight = 0.0
    
    for pipe in Fpipework:
        len_param = get_parameter_value_by_name_AsValueString(pipe, 'Weight')
        lb_param = float (len_param.replace(" lbm", ""))
        if len_param:
            Total_Weight = Total_Weight + lb_param
    # now that results are collected, print the total
    #print("Linear feet of selected pipe(s) is: {}".format(Total_Weight))



    if len(Fhangers) > 0:

        Hanger_Count = 0.0
        
        for hanger in Fhangers:
            hangercount = 1
            Hanger_Count = Hanger_Count + hangercount

    
    pointload = Total_Weight / Hanger_Count

    t = Transaction(doc, 'Write Pointload Info')
    #Start Transaction
    t.Start()

    for whanger in Fhangers:
        numofrods = whanger.GetRodInfo().RodCount
        roundedpointload = round_up(pointload /10) / numofrods
        set_parameter_by_name(whanger,"FP_Pointload", roundedpointload)
        
    #End Transaction
    t.Commit()
else:
    from pyrevit import forms
    forms.alert('At least one fabrication Duct or Hanger was not selected.')