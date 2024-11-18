import Autodesk
from Autodesk.Revit.DB import Transaction, ViewDuplicateOption
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, Button)
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
    #---Convert dialog input into variables
    DuplicateMode = (form.values['DupType'])
    NumberOfCopies = int(form.values['NumofCopies'])

    # Start transaction to perform the duplication in bulk
    t = Transaction(doc, 'Duplicate View(s)')
    t.Start()

    # Dictionary to store views and duplication methods
    duplication_tasks = {}

    for view in selected_views:
        for i in range(NumberOfCopies):
            if DuplicateMode == 'Duplicate':
                if view.ViewType == 'Legend':
                    duplication_tasks[view] = ViewDuplicateOption.WithDetailing
                else:
                    duplication_tasks[view] = ViewDuplicateOption.Duplicate

            elif DuplicateMode == 'Duplicate with Detailing':
                if type(view) == 'Schedule':
                    duplication_tasks[view] = ViewDuplicateOption.Duplicate
                else:
                    duplication_tasks[view] = ViewDuplicateOption.WithDetailing

            elif DuplicateMode == 'Duplicate as a Dependent':
                duplication_tasks[view] = ViewDuplicateOption.AsDependent

    # Process all duplications in the same transaction
    for view, option in duplication_tasks.items():
        for _ in range(NumberOfCopies):
            view.Duplicate(option)

    t.Commit()
except Exception as e:
    forms.alert("An error occurred: {}".format(e), exitscript=True)
    t.RollBack()
    sys.exit()

