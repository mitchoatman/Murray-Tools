import Autodesk
from Autodesk.Revit.DB import Transaction, ViewDuplicateOption
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox, Button)
from pyrevit import forms
import sys

#---Define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

#---Get Selected Views
selected_views = forms.select_views(use_selection=True)
if not selected_views:
    forms.alert("No views selected. Please try again.", exitscript = True)

#---Display dialog to get options
components = [
    Label('Copy Method'),
    ComboBox('DupType', ['Duplicate', 'Duplicate with Detailing', 'Duplicate as a Dependent'], sort=False),
    Label('Number of Copies'),
    TextBox('NumofCopies', '1'),
    Button('Ok')
    ]
form = FlexForm('Duplicate Views', components)
form.show()

try:
    #---Convert dialog input into variable
    DuplicateMode = (form.values['DupType'])
    NumberOfCopies = int(form.values['NumofCopies'])

    t = Transaction(doc, 'Duplicate View(s)')
    t.Start()

    if DuplicateMode == 'Duplicate':
        for view in selected_views:
            if view.ViewType == 'Legend':
                for i in range(NumberOfCopies):
                    view.Duplicate(ViewDuplicateOption.WithDetailing)
            else:
                for i in range(NumberOfCopies):
                    view.Duplicate(ViewDuplicateOption.Duplicate)

    if DuplicateMode == 'Duplicate with Detailing':
        for view in selected_views:
            if type(view) == 'Schedule':
                for i in range(NumberOfCopies):
                    view.Duplicate(ViewDuplicateOption.Duplicate)
            else:
                for i in range(NumberOfCopies):
                    view.Duplicate(ViewDuplicateOption.WithDetailing)

    if DuplicateMode == 'Duplicate as a Dependent':
        for view in selected_views:
            for i in range(NumberOfCopies):
                view.Duplicate(ViewDuplicateOption.AsDependent)
    t.Commit()
except:
        sys.exit()