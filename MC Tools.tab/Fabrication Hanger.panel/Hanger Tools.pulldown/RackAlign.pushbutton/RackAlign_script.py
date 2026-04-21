import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Transaction, ElementId, FilteredElementCollector, BuiltInCategory, BuiltInParameter
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, Separator, Button, CheckBox, Alert
import re

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = int(RevitVersion)

def parse_elevation(input_str):
    """
    Converts various user-input formats into feet as a float.
    Supported formats:
    - 5-6         -> 5 feet 6 inches
    - 5 6         -> 5 feet 6 inches
    - 5.5         -> 5.5 feet
    - 5'-6"       -> 5 feet 6 inches
    - 5' 6"       -> 5 feet 6 inches
    - 5'-6 3/8"   -> 5 feet 6.375 inches
    - 5'-6.125"   -> 5 feet 6.125 inches
    - 0'-0"       -> 0 feet
    - 0'-5 1/2"   -> 0 feet 5.5 inches
    - 0'-0.625"   -> 0 feet 0.625 inches
    - 0'-5.5"     -> 0 feet 5.5 inches
    - 55'-0"      -> 55 feet
    - 18' 6 77/128" -> 18 feet 6.6015625 inches
    """
    input_str = input_str.strip().replace('"', '')  # Remove quotes if present

    # Case 1: Decimal feet (e.g., "5.5")
    if re.match(r"^\d+(\.\d+)?$", input_str):
        return float(input_str)

    # Case 2: Feet-inches format (e.g., "5'-6"", "0'-5 1/2", "55'-0", "18' 6 77/128")
    match = re.match(r"(\d+)['\s\-]*(?:(?:(\d+)\s+(\d+/\d+)|(\d*(?:\.\d+)?))|\d*)?", input_str)
    if match:
        feet = float(match.group(1) or 0)  # Feet part, default to 0 if not present
        inches = 0

        if match.group(2) and match.group(3):  # Handles whole inches + fraction (e.g., "6 77/128")
            whole_inches = float(match.group(2))
            fraction = match.group(3)
            inches = whole_inches + eval(fraction)
        elif match.group(4):  # Handles decimal or whole inches (e.g., "5.5" or "6")
            inches = float(match.group(4))
        elif match.group(2):  # Handles whole inches only (e.g., "6")
            inches = float(match.group(2))

        return feet + (inches / 12)

    raise ValueError("Invalid elevation format. Use 5'-6\", 5-6, 5.5, 0'-5 1/2\", 0'-5.5\", 18' 6 77/128\", etc.")

selected_elements = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

# FUNCTION TO GET PARAMETER VALUE  change "AsDouble()" to "AsString()" to change data type.
def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()
def get_parameter_value(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

if selected_elements and RevitINT > 2022:
    ElevationEstimate = get_parameter_value_by_name(selected_elements[0], 'Lower End Bottom Elevation')
else:
    ElevationEstimate = get_parameter_value_by_name(selected_elements[0], 'Bottom')

# Display dialog
components = [
    CheckBox('TOPmode', 'Align TOP', default=False),
    CheckBox('BTMmode', 'Align BOTTOM', default=True),
    CheckBox('INSMmode', 'Ignore Insulation', default=True),
    Label('Reference Btm Elevation:'),
    TextBox('Elev', default=str(ElevationEstimate)),
    Button('Ok')
]
form = FlexForm('Alignment Method', components)
form.show()

try:
 
    # Convert dialog input into variable
    PRTElevation = target_elevation = parse_elevation(form.values['Elev'])
    TOP = (form.values['TOPmode'])
    BTM = (form.values['BTMmode'])
    INS = (form.values['INSMmode'])

    t = Transaction(doc, "Rack Align")
    t.Start()

    if RevitINT > 2022:
        for elem in selected_elements:
            isfabpart = elem.LookupParameter("Fabrication Service")
            if isfabpart:
                if elem.ItemCustomId in [2041, 866, 40]:
                    if BTM:
                        elem.get_Parameter(BuiltInParameter.MEP_LOWER_BOTTOM_ELEVATION).Set(PRTElevation)
                    if INS and elem.InsulationThickness:
                        INSthickness = get_parameter_value(elem, 'Insulation Thickness')
                        elem.get_Parameter(BuiltInParameter.MEP_LOWER_BOTTOM_ELEVATION).Set(PRTElevation - INSthickness)
                    if TOP:
                        elem.get_Parameter(BuiltInParameter.MEP_LOWER_TOP_ELEVATION).Set(PRTElevation)
                    if INS and elem.InsulationThickness:
                        INSthickness = get_parameter_value(elem, 'Insulation Thickness')
                        elem.get_Parameter(BuiltInParameter.MEP_LOWER_TOP_ELEVATION).Set(PRTElevation + INSthickness)
    else:
        for elem in selected_elements:
            isfabpart = elem.LookupParameter("Fabrication Service")
            if isfabpart:
                if elem.ItemCustomId in [2041, 866, 40]:
                    if BTM:
                        elem.get_Parameter(BuiltInParameter.FABRICATION_BOTTOM_OF_PART).Set(PRTElevation)
                    if INS and elem.InsulationThickness:
                        INSthickness = get_parameter_value(elem, 'Insulation Thickness')
                        elem.get_Parameter(BuiltInParameter.FABRICATION_BOTTOM_OF_PART).Set(PRTElevation - INSthickness)
                    if TOP:
                        elem.get_Parameter(BuiltInParameter.FABRICATION_TOP_OF_PART).Set(PRTElevation)
                    if INS and elem.InsulationThickness:
                        INSthickness = get_parameter_value(elem, 'Insulation Thickness')
                        elem.get_Parameter(BuiltInParameter.FABRICATION_TOP_OF_PART).Set(PRTElevation + INSthickness)
    t.Commit()
except:
    pass