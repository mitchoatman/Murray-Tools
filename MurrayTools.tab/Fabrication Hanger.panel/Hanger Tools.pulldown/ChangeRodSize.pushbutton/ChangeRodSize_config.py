
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FabricationPart, FabricationAncillaryUsage, Transaction, TransactionGroup
from Autodesk.Revit.UI.Selection import *
from rpw.ui.forms import FlexForm, Label, TextBox, Separator, Button
from pyrevit import script
from SharedParam.Add_Parameters import Shared_Params
import os

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def convert_fractions(string):
    # Remove double quotes (") from the input string.
    string = string.replace('"', '')
    # Split the string into a list of tokens.
    tokens = string.split()
    # Initialize variables to keep track of the integer and fractional parts.
    integer_part = 0
    fractional_part = 0.0
    # Iterate over the tokens and convert the mixed number to a float.
    for token in tokens:
        if " " in token:
            # Split the mixed number into integer and fractional parts.
            integer_part_str, fractional_part_str = token.split(" ")
            # Convert the integer part to an integer.
            integer_part += int(integer_part_str)
            # Convert the fractional part to a float.
            fractional_part_str = fractional_part_str.replace('/', '')
            fractional_part += float(fractional_part_str)
        elif "/" in token:
            # If the token is just a fraction, convert it to a float and add it to the fractional part.
            numerator, denominator = token.split("/")
            fractional_part += float(numerator) / float(denominator)
        else:
            # If the token is a standalone number, add it to the integer part.
            integer_part += float(token)
    # Calculate the final result by adding the integer and fractional parts together.
    result = integer_part + fractional_part
    return result

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsString()

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

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
    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_AutoRodSize.txt')

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    if not os.path.exists(filepath):
        with open((filepath), 'w') as the_file:
            line1 = ''
            line2 = ''
            line3 = ''
            line4 = ''
            line5 = ''
            the_file.writelines([line1, line2, line3, line4, line5])  

        # read text file for stored values and show them in dialog
    with open((filepath), 'r') as file:
        lines = file.readlines()
        lines = [line.rstrip() for line in lines]

    if lines:
        # Display dialog with stored values
        components = [
            Label('7/8 Rod / Max Pipe Size:'),
            TextBox('rod_875',lines[0]),
            Label('3/4 Rod / Max Pipe Size:'),
            TextBox('rod_075', lines[1]),
            Label('5/8 Rod / Max Pipe Size:'),
            TextBox('rod_625', lines[2]),
            Label('1/2 Rod / Max Pipe Size'),
            TextBox('rod_050', lines[3]),
            Label('3/8 Rod / Max Pipe Size:'),
            TextBox('rod_375', lines[4]),
            Button('Ok')
            ]
        form = FlexForm('Hanger Rod Sizing', components)
        form.show()
    else:
        # Display dialog without stored values
        components = [
            Label('7/8 Rod / Max Pipe Size:'),
            TextBox('rod_875',''),
            Label('3/4 Rod / Max Pipe Size:'),
            TextBox('rod_075', ''),
            Label('5/8 Rod / Max Pipe Size:'),
            TextBox('rod_625', ''),
            Label('1/2 Rod / Max Pipe Size'),
            TextBox('rod_050', ''),
            Label('3/8 Rod / Max Pipe Size:'),
            TextBox('rod_375', ''),
            Button('Ok')
            ]
        form = FlexForm('Hanger Rod Sizing', components)
        form.show()

    # Convert dialog input into variable

    rod875 = convert_fractions(form.values['rod_875'])
    rod075 = convert_fractions(form.values['rod_075'])
    rod625 = convert_fractions(form.values['rod_625'])
    rod050 = convert_fractions(form.values['rod_050'])
    rod375 = convert_fractions(form.values['rod_375'])

    # write values to text file for future retrieval
    with open((filepath), 'w') as the_file:
        line1 = (str(rod875) + '\n')
        line2 = (str(rod075) + '\n')
        line3 = (str(rod625) + '\n')
        line4 = (str(rod050) + '\n')
        line5 = (str(rod375) + '\n')
        the_file.writelines([line1, line2, line3, line4, line5])

    if rod875 == 0.0 or '':
        rod875 = 0
    if rod075 == 0.0 or '':
        rod075 = 0
    if rod625 == 0.0 or '':
        rod625 = 0
    if rod050 == 0.0 or '':
        rod050 = 0
    if rod375 == 0.0 or '':
        rod375 = 0

    #-----left this info for reference-----
    # if value == 'G - 1-1/4':
        # newrodkit = 70
    # if value == 'F - 1':
        # newrodkit = 67
    # if rod875:
        # newrodkit = 64
    # if rod075:
        # newrodkit = 62
    # if rod625:
        # newrodkit = 31
    # if rod050:
        # newrodkit = 42
    # if rod375:
        # newrodkit = 58
    #-----left this info for reference-----

    pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers")            
    hangers = [doc.GetElement( elId ) for elId in pipesel]

    tg = TransactionGroup(doc, "Change Hanger Rod")
    tg.Start()

    t = Transaction(doc, "Set Hanger Rod")
    t.Start()
    for hanger in hangers:
        hosted_info = hanger.GetHostedInfo().HostId
        try:
            # Get the host element's size
            HostSize = convert_fractions(get_parameter_value_by_name(doc.GetElement(hosted_info), 'Size'))
            if HostSize <= rod875:
                newrodkit = 64
            if HostSize <= rod075:
                newrodkit = 62
            if HostSize <= rod625:
                newrodkit = 31
            if HostSize <= rod050:
                newrodkit = 42
            if HostSize <= rod375:
                newrodkit = 58
            # Set rod size.
            hanger.HangerRodKit = newrodkit
        except:
            output = script.get_output()
            print('{}: {}'.format('Disconnected Hanger', output.linkify(hanger.Id)))
    t.Commit()
    
    t = Transaction(doc, "Update FP Parameter")
    t.Start()
    for x in hangers:
        [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]
    t.Commit()
    
    #End Transaction Group
    tg.Assimilate()
    
except:
    pass



