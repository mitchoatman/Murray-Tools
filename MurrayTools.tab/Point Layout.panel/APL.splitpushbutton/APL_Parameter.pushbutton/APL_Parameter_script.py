
#Imports
from pyrevit import DB, revit, script
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, ElementId
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsDouble
from rpw.ui.forms import FlexForm, Label, TextBox, Button, CheckBox
from SharedParam.Add_Parameters import Shared_Params
import os

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

try:
    OBJselection = uidoc.Selection.PickObjects(ObjectType.Element, 'Select Elements or Finish Button')         
    selection = [doc.GetElement( elId ) for elId in OBJselection]

    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_PointLayout.txt')

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    if not os.path.exists(filepath):
        f = open((filepath), 'w')
        f.write('123')
        f.close()

    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_PointLayout.txt')
    try:
        # read text file for stored values and show them in dialog
        with open((filepath), 'r') as file:
            lines = file.readlines()
            lines = [line.rstrip() for line in lines]
    except:
        with open((filepath), 'w') as the_file:
            line1 = ('pre' + '\n')
            line2 = ('num' + '\n')
            the_file.writelines([line1, line2]) 

    if len(lines) < 2:
        with open((filepath), 'w') as the_file:
            line1 = ('desc' + '\n')
            line2 = ('pre' + '\n')
            the_file.writelines([line1, line2]) 

    # read text file for stored values and show them in dialog
    with open((filepath), 'r') as file:
        lines = file.readlines()
        lines = [line.rstrip() for line in lines]


    # Display dialog
    components = [

        Label('Point Prefix:'),
        TextBox('Pre', lines[1]),
        CheckBox('changepre', '[Enable] the Prefix (selection)', default=False),
        Label('Point Description:'),
        TextBox('Desc', lines[0]),
        CheckBox('changedesc', '[Enable] the Description (selection)', default=False),
        CheckBox('cleanupins', 'Re-format Insert Description (view)', default=False),
        CheckBox('writeslvdims', 'Add Size and Length to Sleeve Description (view)', default=False),
        Button('Ok')
        ]
    form = FlexForm('Update APL Information', components)
    form.show()


    # Convert dialog input into variable
    value = (form.values['Desc']).upper()
    value1 = (form.values['Pre']).upper()
    chkpre = (form.values['changepre'])
    chkdesc = (form.values['changedesc'])
    chkins = (form.values['cleanupins'])
    chkslv = (form.values['writeslvdims'])


    # write values to text file for future retrieval
    with open((filepath), 'w') as the_file:
        line1 = (value + '\n')
        line2 = (value1 + '\n')
        the_file.writelines([line1, line2])



    t = Transaction(doc, 'Modify Point Data')
    #Start Transaction
    t.Start()

    try:
        if chkdesc == True:
            for i in selection:
                 if value:
                    param_exist_ts = i.LookupParameter("TS_Point_Description")
                    if param_exist_ts:
                        #writes data to line number parameterzr
                        set_parameter_by_name(i,"TS_Point_Description", value)
                    else:
                        set_parameter_by_name(i,"PointDescription", value)  
        if chkpre == True:
            for i in selection:
                if value1:
                    param_exist = i.LookupParameter("PointNumber")
                    if param_exist:
                        #writes data to line number parameterzr
                        set_parameter_by_name(i,"PointNumber", value1)

                    param_exist_0 = i.LookupParameter("GTP_PointNumber_0")
                    if param_exist_0:
                        #writes data to line number parameterzr
                        set_parameter_by_name(i,"GTP_PointNumber_0", value1)

                    param_exist_1 = i.LookupParameter("GTP_PointNumber_1")
                    if param_exist_1:
                        #writes data to line number parameterzr
                        set_parameter_by_name(i,"GTP_PointNumber_1", value1)

                    param_exist_2 = i.LookupParameter("GTP_PointNumber_2")
                    if param_exist_2:
                        #writes data to line number parameterzr
                        set_parameter_by_name(i,"GTP_PointNumber_2", value1)

                    param_exist_3 = i.LookupParameter("GTP_PointNumber_3")
                    if param_exist_3:
                        #writes data to line number parameterzr
                        set_parameter_by_name(i,"GTP_PointNumber_3", value1)
        if chkins == True:
            # Create a FilteredElementCollector to get Generic category elements
            generic_models_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_GenericModel)
            # Filter elements by name
            pipe_point_elements = [element for element in generic_models_collector if "Pipe Pt" in element.Name]

            for x in pipe_point_elements:
                originalvalue = get_parameter_value_by_name_AsString(x, 'PointDescription')
                result_string = originalvalue.replace("0' - 0 ", "")
                set_parameter_by_name(x, 'PointDescription', result_string)
        if chkslv == True:
            # Create a FilteredElementCollector to get Pipe Accessory elements
            accessory_models_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_PipeAccessory)
            # Filter elements by name
            accessory_elements = [element for element in accessory_models_collector if "Metal Sleeve" in element.Name or "Plastic Sleeve" in element.Name or "Cast Iron Sleeve" in element.Name]
            for x in accessory_elements:
                slvdiameter = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Pipe Nominal Diameter') * 12)
                slvlength = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Sleeve Length') * 12)
                result_string = 'SLV ' + slvdiameter + ' x ' + slvlength
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            accessory_elements2 = [element for element in accessory_models_collector if "Pipe Riser" in element.Name]
            for x in accessory_elements2:
                slvdiameter = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Diameter') * 12)
                result_string2 = slvdiameter + ' RISER'
                set_parameter_by_name(x, 'TS_Point_Description', result_string2)
    except:
        pass
        print 'Something did not get data, good luck!  Trust but verify...'
    #End Transaction
    t.Commit()
except:
    pass