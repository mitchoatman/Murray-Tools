
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.UI.Selection import ObjectType, PickBoxStyle, Selection

#pyrevit
from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document

selected_views = forms.select_views(title='Select Sheet or Views to Rename')

if selected_views:

    from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)
    components = [Label('Prefix:'),  TextBox('prefix'),
                  Label('Find: (Case Sensitive)'),    TextBox('find'),
                  Label('Replace:'), TextBox('replace'),
                  Label('Suffix:'),  TextBox('suffix'),
                  Separator(),       Button('Rename Views')]

    form = FlexForm('Rename Sheets and Views', components)
    form.show()

    try:
        user_inputs = form.values
        prefix  = user_inputs['prefix']
        find    = user_inputs['find']
        replace = user_inputs['replace']
        suffix  = user_inputs['suffix']

        t = Transaction(doc, 'Rename Sheets and Views')
        t.Start()
        for view in selected_views:
            current_name = view.Name
            new_name = prefix + view.Name.replace(find, replace) + suffix
            for i in range(20):
                try:
                    view.Name = new_name
                    print("{} ->-> {}".format(current_name, new_name))
                    break
                except:
                    new_name += "*"
        t.Commit()
    except:
        print 'User aborted operation, nothing renamed.'
else:
    print 'User did not select any sheets or views.'