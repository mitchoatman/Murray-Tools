from Autodesk.Revit.DB import Transaction
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString, set_parameter_by_name
# pyrevit
from pyrevit import forms, revit


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

selected_elem = revit.get_selection()

if selected_elem:
    from rpw.ui.forms import (FlexForm, Label, TextBox, Button)
    components = [Label('Find: (Case Sensitive)'), TextBox('find'),
                  Label('Replace:'), TextBox('replace'),
                  Button('Rename Spool and Package')]

    form = FlexForm('Rename Spool and Package', components)
    form.show()

    user_inputs = form.values
    find = user_inputs['find']
    replace = user_inputs['replace']

    sname_changes = []
    pname_changes = []

    t = Transaction(doc, 'Rename Spool')
    t.Start()
    for elem in selected_elem:
        if elem.LookupParameter("Fabrication Service"):
            current_sname = get_parameter_value_by_name_AsString(elem, "STRATUS Assembly")
            current_pname = get_parameter_value_by_name_AsString(elem, "STRATUS Package")
            if current_sname:
                new_sname = current_sname.replace(find, replace)
                set_parameter_by_name(elem, "STRATUS Assembly", new_sname)
                sname_changes.append((current_sname, new_sname))
            if current_pname:
                new_pname = current_pname.replace(find, replace)
                set_parameter_by_name(elem, "STRATUS Package", new_pname)
                pname_changes.append((current_pname, new_pname))
    t.Commit()

    # Print all sname changes
    print("STRATUS Assembly changes:")
    for current, new in sname_changes:
        print("{} ->-> {}".format(current, new))

    # Print all pname changes
    print("STRATUS Package changes:")
    for current, new in pname_changes:
        print("{} ->-> {}".format(current, new))

else:
    print('User did not select any elements.')
