
import Autodesk
from Autodesk.Revit.DB import FabricationAncillaryUsage, FabricationPart
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CheckBox
import sys, math
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger, get_parameter_value_by_name_AsDouble
import re

def feet_inches_to_float(feet_inches_str):
    # Regex pattern to match feet, dash, inches, and optional fraction
    pattern = r"(\d+)'[\s-]*(\d+)?\s*(\d+/\d+)?\"?"
    match = re.match(pattern, feet_inches_str)
    
    if match:
        feet = int(match.group(1))  # Feet
        inches = int(match.group(2)) if match.group(2) else 0  # Inches (optional)
        fraction = match.group(3)  # Fraction, e.g., 1/2 (optional)

        # Convert feet and inches to total inches
        total_inches = feet * 12 + inches
        
        # If there's a fraction, convert it to a decimal and add to total inches
        if fraction:
            num, denom = map(int, fraction.split('/'))
            total_inches += num / denom
        
        # Convert total inches to feet
        return total_inches / 12


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

#This writes to fab part custom data field
def set_customdata_by_custid(fabpart, custid, value):
	fabpart.SetPartCustomDataReal(custid, value)

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

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

def set_fp_parameters(element, InsertDepth, TRL):
    # Set the 'FP_Insert Depth' parameter
    set_parameter_by_name(element, 'FP_Insert Depth', InsertDepth)
    # Calculate and set the 'FP_Rod Cut Length' parameter
    rod_cut_length = InsertDepth + TRL
    set_parameter_by_name(element, 'FP_Rod Cut Length', round_to_nearest_quarter(rod_cut_length))

def round_to_nearest_half(value):
    value_in_inches = value * 12
    rounded_value_in_inches = math.ceil(value_in_inches * 2) / 2
    return rounded_value_in_inches / 12

def round_to_nearest_quarter(value):
    value_in_inches = value * 12
    rounded_value_in_inches = math.ceil(value_in_inches * 4) / 4
    return rounded_value_in_inches / 12

pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
CustomISelectionFilter("Pipe Accessories"), "Select Stiffy")            
hangers = [doc.GetElement( elId ) for elId in pipesel]

if len(hangers) > 0:
  
        #Define dialog options and show it
    components = [Label('Rod Elevation:'),
        TextBox('Relev', '3'),
        Button('Ok')]
    form = FlexForm('Rod Elevation (inches)', components)
    form.show()

    DeckThickness = (float(form.values['Relev']) / 12)

    t = Transaction(doc, "Set Insert Depth")
    t.Start()  
    for hanger in hangers:
        hgr_elev = get_parameter_value_by_name_AsDouble(hanger, 'Elevation from Level')
        new_rod_len = DeckThickness - hgr_elev
        try:
            set_parameter_by_name(hanger, 'Rod Elevation', new_rod_len)
        except:
            pass
        try:
            set_parameter_by_name(hanger, 'DIM D', new_rod_len)
        except:
            pass
    t.Commit()
