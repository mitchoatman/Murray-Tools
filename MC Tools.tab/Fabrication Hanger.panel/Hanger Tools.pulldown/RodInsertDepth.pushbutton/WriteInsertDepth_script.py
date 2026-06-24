import Autodesk
from Autodesk.Revit.DB import FabricationAncillaryUsage, FabricationPart, Transaction
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString
from collections import OrderedDict
import clr
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode, HorizontalAlignment
from System.Windows.Controls import StackPanel, Label, ComboBox, TextBox, Button, Orientation
from System.Windows.Media import FontFamily
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
    is_beam_hanger = False
    try:
        beam_hanger_param = get_parameter_value_by_name(hanger, 'FP_Beam Hanger')
        is_beam_hanger = beam_hanger_param == 'Yes'
    except:
        pass

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

    if insert_type == 2:  # BlueBanger-2 MD
        set_parameter_by_name(hanger, 'FP_Insert Type', 'BlueBanger-2 MD - ' + str(deck_thickness * 12))
        if not is_beam_hanger:
            AnciObj = hanger.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250':  # 3/8"
                        InsertDepth = (deck_thickness + (1.625 / 12))
                    elif formatted_elinfo == '0.041667':  # 1/2"
                        InsertDepth = (deck_thickness + (1.25 / 12))
                    elif formatted_elinfo == '0.052083':  # 5/8"
                        InsertDepth = (deck_thickness + (1.875 / 12))
                    elif formatted_elinfo == '0.062500':  # 3/4"
                        InsertDepth = (deck_thickness + (1.125 / 12))
                    elif formatted_elinfo in ['0.072917', '0.083333']:  # 7/8" or 1"
                        InsertDepth = (-4.0 / 12)

    elif insert_type == 3:  # BlueBanger WD
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

    elif insert_type == 4:  # BlueBanger RDI
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

    elif insert_type == 5:  # Dewalt DDI
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

    elif insert_type == 6:  # Dewalt Wood Knocker
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

    elif insert_type == 7:  # Dewalt BangIt+
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

    elif insert_type == 8:  # Hilti KCM-WF
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

    elif insert_type == 9:  # Mason West PAL_CIP WD
        set_parameter_by_name(hanger, 'FP_Insert Type', 'Mason West PAL_CIP WD')
        if not is_beam_hanger:
            AnciObj = hanger.GetPartAncillaryUsage()
            for n in AnciObj:
                AnciDiam.append(n.AncillaryWidthOrDiameter)
                for elinfo in AnciDiam:
                    formatted_elinfo = "{:.6f}".format(elinfo)
                    if formatted_elinfo == '0.031250':  # 3/8"
                        InsertDepth = (2.375 / 12)
                    elif formatted_elinfo == '0.041667':  # 1/2"
                        InsertDepth = (1.75 / 12)
                    elif formatted_elinfo == '0.052083':  # 5/8"
                        InsertDepth = (1.0 / 12)
                    elif formatted_elinfo == '0.062500':  # 3/4"
                        InsertDepth = (0.875 / 12)
                    elif formatted_elinfo in ['0.072917', '0.083333']:  # 7/8" or 1"
                        InsertDepth = (1.0 / 12)

    elif insert_type == 10:  # Custom Cut/Extended Rod
        set_parameter_by_name(hanger, 'FP_Insert Type', 'Custom Cut')
        if not is_beam_hanger:
            InsertDepth = deck_thickness

    # Only set custom data and parameters if InsertDepth is not None
    if InsertDepth is not None:
        set_customdata_by_custid(hanger, 9, InsertDepth)
        set_fp_parameters(hanger, InsertDepth, rla, is_beam_hanger)

class InsertTypeDialog(Window):
    def __init__(self):
        self.Title = "Insert Type"
        self.Width = 340
        self.Height = 220
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.values = {}

        self.insert_options = OrderedDict([
            ('(01) BlueBanger MD', 1),
            ('(02) BlueBanger-2 MD', 2),
            ('(03) BlueBanger WD', 3),
            ('(04) BlueBanger RDI', 4),
            ('(05) Dewalt DDI', 5),
            ('(06) Dewalt Wood Knocker', 6),
            ('(07) Dewalt BangIt+', 7),
            ('(08) Hilti KCM-WF', 8),
            ('(09) Mason West PAL_CIP WD', 9),
            ('(10) Extend or Cut Rod', 10)
        ])

        stack = StackPanel()
        stack.Orientation = Orientation.Vertical
        stack.Margin = Thickness(10)

        label_insert = Label()
        label_insert.Content = 'Insert Type:'
        label_insert.FontSize = 12
        label_insert.FontFamily = FontFamily("Arial")
        stack.Children.Add(label_insert)

        self.combobox_insert = ComboBox()
        self.combobox_insert.Width = 300
        self.combobox_insert.Height = 24
        self.combobox_insert.FontSize = 12
        self.combobox_insert.FontFamily = FontFamily("Arial")
        self.combobox_insert.HorizontalAlignment = HorizontalAlignment.Left
        self.combobox_insert.Margin = Thickness(0, 0, 0, 10)
        for key in self.insert_options.keys():
            self.combobox_insert.Items.Add(key)
        self.combobox_insert.SelectedIndex = 0
        stack.Children.Add(self.combobox_insert)

        label_deck = Label()
        label_deck.Content = 'Deck Thickness(Decimal In.) \n(Pos Num = Ext / Neg Num = Cut):'
        label_deck.FontSize = 12
        label_deck.FontFamily = FontFamily("Arial")
        stack.Children.Add(label_deck)

        self.textbox_deck = TextBox()
        self.textbox_deck.Width = 300
        self.textbox_deck.Height = 20
        self.textbox_deck.FontSize = 12
        self.textbox_deck.FontFamily = FontFamily("Arial")
        self.textbox_deck.Text = '3'
        self.textbox_deck.HorizontalAlignment = HorizontalAlignment.Left
        self.textbox_deck.Margin = Thickness(0, 0, 0, 10)
        stack.Children.Add(self.textbox_deck)

        self.button_ok = Button()
        self.button_ok.Content = 'Ok'
        self.button_ok.Width = 74
        self.button_ok.Height = 25
        self.button_ok.HorizontalAlignment = HorizontalAlignment.Center
        self.button_ok.Click += self.ok_button_clicked
        stack.Children.Add(self.button_ok)

        self.Content = stack

    def ok_button_clicked(self, sender, event):
        selected_key = self.combobox_insert.SelectedItem
        self.values = {
            'Insert': self.insert_options[selected_key],
            'Deck': self.textbox_deck.Text
        }
        self.DialogResult = True
        self.Close()

# Main execution
pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers")
hangers = [doc.GetElement(elId) for elId in pipesel]

if len(hangers) > 0:
    form = InsertTypeDialog()
    result = form.ShowDialog()

    if not result or not form.values or 'Insert' not in form.values:
        sys.exit()

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
        RLA = 0.0
        RLB = 0.0
        TrapWidth = None
        TrapExtn = None
        
        for dta in ItmDims:
            try:
                if dta.Name == 'Length A':
                    RLA = e.GetDimensionValue(dta) or 0.0
                elif dta.Name == 'Length B':
                    RLB = e.GetDimensionValue(dta) or 0.0
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
                set_parameter_by_name(e, 'FP_Rod Length A', RLA)
                set_parameter_by_name(e, 'FP_Rod Length B', RLB)
        except Exception:
            pass

        # Process insert type for this hanger
        rod_size = get_parameter_value_by_name(e, 'FP_Rod Size')
        if rod_size is not None:
            process_hanger_insert(e, InsertType, DeckThickness, RLA)

    t.Commit()