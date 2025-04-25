import Autodesk
from Autodesk.Revit.DB import FabricationAncillaryUsage, FabricationPart, Transaction
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, Separator, Button, CheckBox
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString
import sys, math

Shared_Params()
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Existing helper functions remain unchanged
def set_customdata_by_custid(fabpart, custid, value):
    fabpart.SetPartCustomDataReal(custid, value)

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsString() if element.LookupParameter(parameterName).StorageType == Autodesk.Revit.DB.StorageType.String else element.LookupParameter(parameterName).AsDouble()

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return True

def round_to_nearest_half(value):
    value_in_inches = value * 12
    rounded_value_in_inches = math.ceil(value_in_inches * 2) / 2
    return rounded_value_in_inches / 12

def round_to_nearest_quarter(value):
    value_in_inches = value * 12
    rounded_value_in_inches = math.ceil(value_in_inches * 4) / 4
    return rounded_value_in_inches / 12

def set_fp_parameters(element, InsertDepth, TRL, is_beam_hanger):
    set_parameter_by_name(element, 'FP_Insert Depth', InsertDepth)
    # Only add InsertDepth to rod length if not a beam hanger
    rod_cut_length = TRL if is_beam_hanger else (InsertDepth + TRL)
    set_parameter_by_name(element, 'FP_Rod Cut Length', round_to_nearest_half(rod_cut_length))

# Updated function to process each hanger individually with beam hanger check
def process_hanger_insert(hanger, insert_type, deck_thickness, rla):
    AnciDiam = list()
    # Check if this is a beam hanger
    is_beam_hanger = False
    try:
        beam_hanger_param = get_parameter_value_by_name(hanger, 'FP_Beam Hanger')
        is_beam_hanger = beam_hanger_param == 'Yes'
    except:
        pass  # If parameter doesn't exist or fails, assume it's not a beam hanger

    # If it's a beam hanger, set InsertDepth to 0
    InsertDepth = 0 if is_beam_hanger else None

    if insert_type == 1:  # BlueBanger MD
        set_parameter_by_name(hanger, 'FP_Insert Type', 'BlueBanger MD - ' + str(deck_thickness * 12))
        if not is_beam_hanger:
            AnciObj = hanger.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250':  # 3/8"
                        InsertDepth = (deck_thickness + (1.25 / 12))
                    elif formatted_elinfo == '0.041667':  # 1/2"
                        InsertDepth = (deck_thickness + (0.75 / 12))
                    elif formatted_elinfo == '0.052083':  # 5/8"
                        InsertDepth = (deck_thickness + (0.75 / 12))
                    elif formatted_elinfo == '0.062500':  # 3/4"
                        InsertDepth = (deck_thickness + (1.0 / 12))
                    elif formatted_elinfo in ['0.072917', '0.083333']:  # 7/8" or 1"
                        InsertDepth = (-4.0 / 12)

    elif insert_type == 2:  # BlueBanger WD
        set_parameter_by_name(hanger, 'FP_Insert Type', 'BlueBanger WD')
        if not is_beam_hanger:
            AnciObj = hanger.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250':  # 3/8"
                        InsertDepth = (1.75 / 12)
                    elif formatted_elinfo == '0.041667':  # 1/2"
                        InsertDepth = (1.25 / 12)
                    elif formatted_elinfo == '0.052083':  # 5/8"
                        InsertDepth = (1.75 / 12)
                    elif formatted_elinfo == '0.062500':  # 3/4"
                        InsertDepth = (1.25 / 12)
                    elif formatted_elinfo in ['0.072917', '0.083333']:  # 7/8" or 1"
                        InsertDepth = (-4.0 / 12)

    elif insert_type == 3:  # BlueBanger RDI
        set_parameter_by_name(hanger, 'FP_Insert Type', 'BlueBanger RDI')
        if not is_beam_hanger:
            AnciObj = hanger.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250':  # 3/8"
                        InsertDepth = (deck_thickness + (-0.75 / 12))
                    elif formatted_elinfo == '0.041667':  # 1/2"
                        InsertDepth = (deck_thickness + (-1.25 / 12))
                    elif float(formatted_elinfo) > 0.052083:  # Bigger than 1/2"
                        print('Some rod sizes bigger than insert can accept! \n Depth not written for hanger.')
                        InsertDepth = 0

    elif insert_type == 4:  # Dewalt DDI
        set_parameter_by_name(hanger, 'FP_Insert Type', 'Dewalt DDI')
        if not is_beam_hanger:
            AnciObj = hanger.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250':  # 3/8"
                        InsertDepth = (deck_thickness + (-6.25 / 12))
                    elif formatted_elinfo == '0.041667':  # 1/2"
                        InsertDepth = (deck_thickness + (-6.0 / 12))
                    elif formatted_elinfo == '0.052083':  # 5/8"
                        InsertDepth = (deck_thickness + (-5.625 / 12))
                    elif formatted_elinfo in ['0.062500', '0.072917']:  # 3/4" or 7/8"
                        InsertDepth = (deck_thickness + (-5.375 / 12))
                    elif formatted_elinfo == '0.083333':  # 1"
                        InsertDepth = (-4.0 / 12)

    elif insert_type == 5:  # Dewalt Wood Knocker
        set_parameter_by_name(hanger, 'FP_Insert Type', 'Dewalt Wood Knocker')
        if not is_beam_hanger:
            AnciObj = hanger.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250':  # 3/8"
                        InsertDepth = (0.875 / 12)
                    elif formatted_elinfo in ['0.041667', '0.052083', '0.062500']:  # 1/2", 5/8", 3/4"
                        InsertDepth = (0.375 / 12)
                    elif formatted_elinfo in ['0.072917', '0.083333']:  # 7/8" or 1"
                        InsertDepth = (-4.0 / 12)

    elif insert_type == 6:  # Dewalt BangIt+
        set_parameter_by_name(hanger, 'FP_Insert Type', 'Dewalt BangIt+')
        if not is_beam_hanger:
            AnciObj = hanger.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250':  # 3/8"
                        InsertDepth = (0.375 / 12)
                    elif formatted_elinfo == '0.041667':  # 1/2"
                        InsertDepth = (0.500 / 12)
                    elif formatted_elinfo == '0.052083':  # 5/8"
                        InsertDepth = (0.625 / 12)
                    elif formatted_elinfo == '0.062500':  # 3/4"
                        InsertDepth = (0.750 / 12)
                    elif formatted_elinfo in ['0.072917', '0.083333']:  # 7/8" or 1"
                        InsertDepth = (-4.0 / 12)

    elif insert_type == 7:  # Hilti KCM-WF
        set_parameter_by_name(hanger, 'FP_Insert Type', 'Hilti KCM-WF')
        if not is_beam_hanger:
            AnciObj = hanger.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250':  # 3/8"
                        InsertDepth = (1.9 / 12)
                    elif formatted_elinfo == '0.041667':  # 1/2"
                        InsertDepth = (1.5 / 12)
                    elif formatted_elinfo == '0.052083':  # 5/8"
                        InsertDepth = (0.9 / 12)
                    elif formatted_elinfo == '0.062500':  # 3/4"
                        InsertDepth = (1.5 / 12)
                    elif formatted_elinfo in ['0.072917', '0.083333']:  # 7/8" or 1"
                        InsertDepth = (-4.0 / 12)

    elif insert_type == 8:  # Custom Cut/Extended Rod
        set_parameter_by_name(hanger, 'FP_Insert Type', 'User Custom')
        if not is_beam_hanger:
            InsertDepth = deck_thickness

    # Only set custom data and parameters if InsertDepth is not None
    if InsertDepth is not None:
        set_customdata_by_custid(hanger, 9, InsertDepth)
        set_fp_parameters(hanger, InsertDepth, rla, is_beam_hanger)

# Main execution
pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers")
hangers = [doc.GetElement(elId) for elId in pipesel]

if len(hangers) > 0:
    # Define dialog options and show it
    components = [
        Label('Insert Type:'),
        ComboBox('Insert', {
            '(1) BlueBanger MD': 1,
            '(2) BlueBanger WD': 2,
            '(3) BlueBanger RDI': 3,
            '(4) Dewalt DDI': 4,
            '(5) Dewalt Wood Knocker': 5,
            '(6) Dewalt BangIt+': 6,
            '(7) Hilti KCM-WF': 7,
            '(8) Extend or Cut Rod': 8
        }),
        Label('Deck Thickness(Decimal In.) | (Num = +Ext or -Cut):'),
        TextBox('Deck', '3'),
        Button('Ok')
    ]
    form = FlexForm('Insert Type', components)
    form.show()

    InsertType = form.values['Insert']
    DeckThickness = float(form.values['Deck']) / 12

    t = Transaction(doc, "Set Insert Depth")
    t.Start()
    
    for e in hangers:
        # Skip elements named 'Seismic LRD'
        try:
            element_name = get_parameter_value_by_name_AsString(e, 'Family')
            if element_name == 'Seismic LRD':
                continue
        except:
            pass

        try:
            [set_parameter_by_name(e, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in e.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]
            [set_parameter_by_name(e, 'FP_Product Entry', get_parameter_value_by_name_AsString(e, 'Product Entry')) if e.LookupParameter('Product Entry') else set_parameter_by_name(e, 'FP_Product Entry', get_parameter_value_by_name_AsString(e, 'Size'))]
        except:
            pass
        # Get hanger dimensions
        ItmDims = e.GetDimensions()
        RLA = None
        RLB = None
        TrapWidth = None
        TrapExtn = None
        
        for dta in ItmDims:
            try:
                if dta.Name == 'Length A':
                    RLA = e.GetDimensionValue(dta)
                elif dta.Name == 'Length B':
                    RLB = e.GetDimensionValue(dta)
                elif dta.Name == 'Width':
                    TrapWidth = e.GetDimensionValue(dta)
                elif dta.Name == 'Bearer Extn':
                    TrapExtn = e.GetDimensionValue(dta)
            except Exception:
                pass

        # Set basic parameters
        try:
            if TrapWidth is not None and TrapExtn is not None:
                BearerLength = TrapWidth + TrapExtn + TrapExtn
                set_parameter_by_name(e, 'FP_Bearer Length', BearerLength)
        except Exception:
            pass

        try:
            if RLA is not None:
                set_parameter_by_name(e, 'FP_Rod Length', RLA)
                set_parameter_by_name(e, 'FP_Rod Length A', RLA)
            if RLB is not None:
                set_parameter_by_name(e, 'FP_Rod Length B', RLB)
        except Exception:
            pass

        # Process insert type for this hanger
        rod_size = get_parameter_value_by_name(e, 'FP_Rod Size')
        if rod_size is not None:  # Only process if we have a rod size
            process_hanger_insert(e, InsertType, DeckThickness, RLA)

    t.Commit()