
import Autodesk
from Autodesk.Revit.DB import FabricationAncillaryUsage, FabricationPart
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CheckBox
from SharedParam.Add_Parameters import Shared_Params
import sys

Shared_Params()
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

pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers")            
hangers = [doc.GetElement( elId ) for elId in pipesel]

if len(hangers) > 0:
    
        #Define dialog options and show it
    components = [Label('Insert Type:'),
        ComboBox('Insert', {'(1) BlueBanger MD': 1, '(2) BlueBanger WD': 2, '(3) BlueBanger RDI': 3, '(4) Dewalt DDI': 4, '(5) Dewalt Wood Knocker': 5, '(6) Extend or Cut Rod': 6}),
        Label('Deck Thickness(Decimal In.) | (Num = +Ext or -Cut):'),
        TextBox('Deck', '3'),
        Button('Ok')]
    form = FlexForm('Insert Type', components)
    form.show()

    InsertType = (form.values['Insert'])
    DeckThickness = (float(form.values['Deck']) / 12)
    
    AnciDiam = list()
    AnciType = list()

    t = Transaction(doc, "Extend BlueBanger MD Rod")
    t.Start() 
    if InsertType == 1:  #BlueBanger MD
        for e in hangers:
            set_parameter_by_name(e, 'FP_Insert Type', 'BlueBanger MD - ' + str(DeckThickness * 12))
            AnciObj = e.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                #print AnciDiam
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250': #3/8"
                        InsertDepth = (DeckThickness + (1.25 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.041667': #1/2"
                        InsertDepth = (DeckThickness + (0.75 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.052083': #5/8"
                        InsertDepth = (DeckThickness + (0.75 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.062500': #3/4"
                        InsertDepth = (DeckThickness + (1.0 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.072917': #7/8"
                        InsertDepth = (-4.0 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.083333': #1"
                        InsertDepth = (-4.0 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)

    if InsertType == 2:  #BlueBanger WD
        for e in hangers:
            set_parameter_by_name(e, 'FP_Insert Type', 'BlueBanger WD')
            AnciObj = e.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                #print AnciDiam
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250': #3/8"
                        InsertDepth = (1.75 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.041667': #1/2"
                        InsertDepth = (1.25 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.052083': #5/8"
                        InsertDepth = (1.75 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.062500': #3/4"
                        InsertDepth = (1.25 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.072917': #7/8"
                        InsertDepth = (-4.0 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.083333': #1"
                        InsertDepth = (-4.0 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)

    if InsertType == 3:  #BlueBanger RDI
        for e in hangers:
            set_parameter_by_name(e, 'FP_Insert Type', 'BlueBanger RDI')
            AnciObj = e.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    #print formatted_elinfo
                    if formatted_elinfo == '0.031250': #3/8"
                        InsertDepth = (DeckThickness + (-0.75 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.041667': #1/2"
                        InsertDepth = (DeckThickness + (-1.25 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo > '0.052083': #Bigger than 1/2"
                        print 'Some rod sizes bigger than insert can accept! \n Depth not written.'

    if InsertType == 4:  #Dewalt DDI
        for e in hangers:
            set_parameter_by_name(e, 'FP_Insert Type', 'Dewalt DDI')
            AnciObj = e.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                #print AnciDiam
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    #print formatted_elinfo
                    if formatted_elinfo == '0.031250': #3/8"
                        InsertDepth = (DeckThickness + (-6.25 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.041667': #1/2"
                        InsertDepth = (DeckThickness + (-6.0 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.052083': #5/8"
                        InsertDepth = (DeckThickness + (-5.625 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.062500': #3/4"
                        InsertDepth = (DeckThickness + (-5.375 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.072917': #7/8"
                        InsertDepth = (DeckThickness + (-5.375 / 12))
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.083333': #1"
                        InsertDepth = (-4.0 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)

    if InsertType == 5:  #Dewalt Wood Knocker
        for e in hangers:
            set_parameter_by_name(e, 'FP_Insert Type', 'Dewalt Wood Knocker')
            AnciObj = e.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                #print AnciDiam
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250': #3/8"
                        InsertDepth = (0.875 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.041667': #1/2"
                        InsertDepth = (0.375 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.052083': #5/8"
                        InsertDepth = (0.375/ 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.062500': #3/4"
                        InsertDepth = (0.375 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.072917': #7/8"
                        InsertDepth = (0.375 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)
                    if formatted_elinfo == '0.083333': #1"
                        InsertDepth = (-4.0 / 12)
                        set_customdata_by_custid(e, 9, InsertDepth)
                        set_parameter_by_name(e, 'FP_Insert Depth', InsertDepth)

    if InsertType == 6:  #Custom Cut/Extended Rod
        for e in hangers:
            set_parameter_by_name(e, 'FP_Insert Type', 'User Custom')
            set_customdata_by_custid(e, 9, DeckThickness)
            set_parameter_by_name(e, 'FP_Insert Depth', DeckThickness)

    t.Commit()
