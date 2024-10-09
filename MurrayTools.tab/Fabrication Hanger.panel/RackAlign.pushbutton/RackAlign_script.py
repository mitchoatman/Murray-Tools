import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Transaction, ElementId, FilteredElementCollector, BuiltInCategory, BuiltInParameter
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, Separator, Button, CheckBox, Alert

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = int(RevitVersion)


selected_elements = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
try:
    #FUNCTION TO GET PARAMETER VALUE  change "AsDouble()" to "AsString()" to change data type.
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
        Label('Reference Btm Elevation ' + '[' + str(ElevationEstimate) + ']:'),
        Label('Elevation use format [FT-IN]:'),
        TextBox('Elev', ''),
        Button('Ok')
        ]
    form = FlexForm('Alignment Method', components)
    form.show()

    # Convert dialog input into variable
    if '-' in (form.values['Elev']):
        InputFT = float((form.values['Elev']).split("-", 1)[0])
        InputIN = (float((form.values['Elev']).split("-", 1)[1]) / 12)
        PRTElevation = InputFT + InputIN
        TOP = (form.values['TOPmode'])
        BTM = (form.values['BTMmode'])
        INS = (form.values['INSMmode'])
    else:
        Alert("You didn't enter the elevation using format [FT-IN]", title="Rack Align Error", header="Elevation Units", exit=True)

    t = Transaction(doc, "Rack Align")
    t.Start()
    if RevitINT > 2022:
        if BTM:
            for elem in selected_elements:
                if elem.ItemCustomId == 2041:
                    elem.get_Parameter(BuiltInParameter.MEP_LOWER_BOTTOM_ELEVATION).Set(PRTElevation)
            if INS:
                for elem in selected_elements:
                    if elem.ItemCustomId == 2041:
                        if elem.InsulationThickness:
                            INSthickness = get_parameter_value(elem, 'Insulation Thickness')
                            elem.get_Parameter(BuiltInParameter.MEP_LOWER_BOTTOM_ELEVATION).Set(PRTElevation - INSthickness)
        if TOP:
            for elem in selected_elements:
                if elem.ItemCustomId == 2041:
                    elem.get_Parameter(BuiltInParameter.MEP_LOWER_TOP_ELEVATION).Set(PRTElevation)
            if INS:
                for elem in selected_elements:
                    if elem.ItemCustomId == 2041:
                        if elem.InsulationThickness:
                            INSthickness = get_parameter_value(elem, 'Insulation Thickness')
                            elem.get_Parameter(BuiltInParameter.MEP_LOWER_TOP_ELEVATION).Set(PRTElevation + INSthickness)
    else:
        if BTM:
            for elem in selected_elements:
                if elem.ItemCustomId == 2041:
                    elem.get_Parameter(BuiltInParameter.FABRICATION_BOTTOM_OF_PART).Set(PRTElevation)
            if INS:
                for elem in selected_elements:
                    if elem.ItemCustomId == 2041:
                        if elem.InsulationThickness:
                            INSthickness = get_parameter_value(elem, 'Insulation Thickness')
                            elem.get_Parameter(BuiltInParameter.FABRICATION_BOTTOM_OF_PART).Set(PRTElevation - INSthickness)
        if TOP:
            for elem in selected_elements:
                if elem.ItemCustomId == 2041:
                    elem.get_Parameter(BuiltInParameter.FABRICATION_TOP_OF_PART).Set(PRTElevation)
            if INS:
                for elem in selected_elements:
                    if elem.ItemCustomId == 2041:
                        if elem.InsulationThickness:
                            INSthickness = get_parameter_value(elem, 'Insulation Thickness')
                            elem.get_Parameter(BuiltInParameter.FABRICATION_TOP_OF_PART).Set(PRTElevation + INSthickness)

    t.Commit()
except:
    pass