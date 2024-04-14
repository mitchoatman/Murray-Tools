
from Autodesk.Revit.DB import FabricationPart
from Autodesk.Revit.DB import FabricationAncillaryUsage
from Autodesk.Revit import DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CheckBox)
from rpw.ui.forms import Alert
import sys

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)
def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()
class MySelectionFilter(ISelectionFilter):
    def __init__(self):
        pass
    def AllowElement(self, element):
        if element.Category.Name == "MEP Fabrication Hangers":
            return True
        else:
            return False
    def AllowReference(self, element):
        return False
selection_filter = MySelectionFilter()
hangers = uidoc.Selection.PickElementsByRectangle(selection_filter)

if len(hangers) > 0:
    
        #Define dialog options and show it
    components = [Label('Insert Type:'),
        ComboBox('Insert', {'(1) BlueBanger MD': 1, '(2) BlueBanger WD': 2, '(3) BlueBanger RDI': 3, '(4) Dewalt DDI': 4, '(5) Dewalt Wood Knocker': 5, '(6) Extend or Cut Rod': 6}),
        Label('Deck Thickness (Decimal Inches):'),
        TextBox('Deck', '3'),
        Button('Ok')]
    form = FlexForm('Insert Type', components)
    form.show()
    try:
        InsertType = (form.values['Insert'])
        DeckThickness = (float(form.values['Deck']) / 12)
        
        AnciObj = list()
        AnciDiam = list()
        AnciType = list()
        ItmList1 = list()   

        t = Transaction(doc, "Write Hanger Data")
        t.Start()
        for e in hangers:
            AnciObj = e.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    if (elinfo) > 0:
                        set_parameter_by_name(e, 'FP_Rod Size', elinfo)
        t.Commit()
        
        if InsertType == 1:  #BlueBanger MD
            try:
                t = Transaction(doc, "Extend BlueBanger MD Rod")
                t.Start()
                for e in hangers:
                    rodsize = get_parameter_value_by_name(e, 'FP_Rod Size')[-4:].strip('"')
                    rodattached = e.GetRodInfo().IsAttachedToStructure
                    if rodattached == True:
                        if  rodsize == '3/8': # 3/8" ROD
                            STName = e.GetRodInfo().RodCount
                            hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                            STName1 = e.GetRodInfo()
                            for n in range(STName):
                                STName1.SetRodStructureExtension(n, (DeckThickness + (1.25 / 12)))
                        if  rodsize == '1/2': # 1/2" ROD
                            STName = e.GetRodInfo().RodCount
                            ItmList1.append(STName)
                            hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                            STName1 = e.GetRodInfo()
                            for n in range(STName):
                                STName1.SetRodStructureExtension(n, (DeckThickness + (0.75 / 12)))
                        if  rodsize == '5/8': # 5/8 ROD
                            STName = e.GetRodInfo().RodCount
                            ItmList1.append(STName)
                            hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                            STName1 = e.GetRodInfo()
                            for n in range(STName):
                                STName1.SetRodStructureExtension(n, (DeckThickness + (0.75 / 12)))
                        if  rodsize == '3/4': # 3/4" ROD
                            STName = e.GetRodInfo().RodCount
                            ItmList1.append(STName)
                            hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                            STName1 = e.GetRodInfo()
                            for n in range(STName):
                                STName1.SetRodStructureExtension(n, (DeckThickness + (1.0 / 12)))
                        if  rodsize == '7/8': # 7/8" ROD
                            STName = e.GetRodInfo().RodCount
                            ItmList1.append(STName)
                            hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                            STName1 = e.GetRodInfo()
                            for n in range(STName):
                                STName1.SetRodStructureExtension(n, (-4.0 / 12))
                        if  rodsize == '- 1': # 1" ROD
                            STName = e.GetRodInfo().RodCount
                            ItmList1.append(STName)
                            hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                            STName1 = e.GetRodInfo()
                            for n in range(STName):
                                STName1.SetRodStructureExtension(n, (-4.0 / 12))
                    else:
                        Alert('Attach all rods to structure before extending rods', title="Insert Extension", header="One or Many Rods Not Attached")
                t.Commit()
            except:
                sys.exit() 
#-----------------------------------------------------------------------#
        if InsertType == 2:  #BlueBanger WD
            try:
                t = Transaction(doc, "Extend BlueBanger WD Rod")
                t.Start()
                if InsertType == 1:  #BlueBanger MD
                    for e in hangers:
                        rodsize = get_parameter_value_by_name(e, 'FP_Rod Size')[-4:].strip('"')
                        rodattached = e.GetRodInfo().IsAttachedToStructure
                        if rodattached == True:
                            if  rodsize == '3/8': # 3/8" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (1.75 / 12))
                            if  rodsize == '1/2': # 1/2" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (1.25 / 12))
                            if  rodsize == '5/8': # 5/8 ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (1.75 / 12))
                            if  rodsize == '3/4': # 3/4" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (1.25 / 12))
                            if  rodsize == '7/8': # 7/8" and larger ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (-4.0 / 12))
                    else:
                        Alert('Attach all rods to structure before extending rods', title="Insert Extension", header="One or Many Rods Not Attached")
                t.Commit()
            except:
                sys.exit() 
#-----------------------------------------------------------------------#
        if InsertType == 3:  #BlueBanger RDI
            try:
                t = Transaction(doc, "Extend BlueBanger RDI Rod")
                t.Start()
                for e in hangers:
                    AnciObj = e.GetPartAncillaryUsage()
                    rodattached = e.GetRodInfo().IsAttachedToStructure
                    if rodattached == True:
                        for n in AnciObj:
                            AnciDiam.append(n.AncillaryWidthOrDiameter)
                            if (0.031 < max(AnciDiam) < 0.032): # 3/8" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (DeckThickness + (-0.75 / 12)))
                            if (0.041 < max(AnciDiam) < 0.042): # 1/2" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (DeckThickness + (-1.25 / 12)))
                    else:
                        Alert('Attach all rods to structure before extending rods', title="Insert Extension", header="One or Many Rods Not Attached")

                t.Commit()
            except:
                sys.exit() 
#-----------------------------------------------------------------------#
        if InsertType == 4:  #Dewalt DDI
            try:
                t = Transaction(doc, "Extend Dewalt DDI Rod")
                t.Start()

                for e in hangers:
                    AnciObj = e.GetPartAncillaryUsage()
                    rodattached = e.GetRodInfo().IsAttachedToStructure
                    if rodattached == True:
                        for n in AnciObj:
                            AnciDiam.append(n.AncillaryWidthOrDiameter)
                            if (0.031 < max(AnciDiam) < 0.032): # 3/8" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (DeckThickness + (-6.25 / 12)))
                            if (0.041 < max(AnciDiam) < 0.042): # 1/2" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (DeckThickness + (-6.0 / 12)))
                            if (0.052 < max(AnciDiam) < 0.053): # 5/8 ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (DeckThickness + (-5.625 / 12)))
                            if (0.062 < max(AnciDiam) < 0.063): # 3/4" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (DeckThickness + (-5.375 / 12)))
                            if (0.072 < max(AnciDiam) < 0.090): # 7/8" and 1" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (DeckThickness + (-5.375 / 12)))
                    else:
                        Alert('Attach all rods to structure before extending rods', title="Insert Extension", header="One or Many Rods Not Attached")
                t.Commit()
            except:
                sys.exit() 
#-----------------------------------------------------------------------#
        if InsertType == 5:  #Dewalt Wood Knocker
            try:
                t = Transaction(doc, "Extend Dewalt Wood Knocker Rod")
                t.Start()

                for e in hangers:
                    AnciObj = e.GetPartAncillaryUsage()
                    rodattached = e.GetRodInfo().IsAttachedToStructure
                    if rodattached == True:
                        for n in AnciObj:
                            AnciDiam.append(n.AncillaryWidthOrDiameter)
                            if (0.031 < max(AnciDiam) < 0.032): # 3/8" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (1.75 / 12))
                            if (0.041 < max(AnciDiam) < 0.042): # 1/2" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (1.25 / 12))
                            if (0.052 < max(AnciDiam) < 0.053): # 5/8 ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (1.75 / 12))
                            if (0.062 < max(AnciDiam) < 0.063): # 3/4" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (1.25 / 12))
                            if (0.072 < max(AnciDiam) < 0.090): # 7/8" and 1" ROD
                                STName = e.GetRodInfo().RodCount
                                ItmList1.append(STName)
                                hgrhost = e.GetRodInfo().CanRodsBeHosted = True
                                STName1 = e.GetRodInfo()
                                for n in range(STName):
                                    STName1.SetRodStructureExtension(n, (-4.0 / 12))
                    else:
                        Alert('Attach all rods to structure before extending rods', title="Insert Extension", header="One or Many Rods Not Attached")
                t.Commit()
            except:
                sys.exit() 
#-----------------------------------------------------------------------#
        if InsertType == 6:  #Extend or Cut Rod
            try:
                t = Transaction(doc, "Extend or Cut Rod")
                t.Start()

                for e in hangers:
                    rodattached = e.GetRodInfo().IsAttachedToStructure
                    #if rodattached == True:
                    STName = e.GetRodInfo().RodCount
                    STName1 = e.GetRodInfo()
                    for n in range(STName):
                        STName1.SetRodStructureExtension(n, (DeckThickness))
                    #else:
                    #    Alert('Attach all rods to structure before extending rods', title="Insert Extension", header="One or Many Rods Not Attached")

                t.Commit()
            except:
                sys.exit() 
    except:
        sys.exit()
else:
    Alert('You must select at least one element', title="Rod Extension", header="No Hangers Selected")

