from Autodesk.Revit.DB import Transaction, FilteredElementCollector, ParameterElement
from pyrevit import DB, revit, forms
from System import Guid
import os

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_RemoveParameters.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    f = open((filepath), 'w')
    f.write('FP_')
    f.close()

f = open((filepath), 'r')
PrevInput = f.read()
f.close()

# This displays dialog
value = forms.ask_for_string(default=PrevInput, prompt='Parameter *Begins With* : (Case Sensitive)\nFP_\nTS_\nSTRATUS\nMultiple comma separated values can be entered', title='Remove Parameters')

if value:

    f = open((filepath), 'w')
    f.write(value)
    f.close()

    # Retrieve all parameters in the document
    params = FilteredElementCollector(doc).OfClass(ParameterElement)
    filteredparams = []

    # Split the input value into multiple entries using comma as the separator
    param_names = [name.strip() for name in value.split(',')]

    for param in params:
        if param.Name.startswith(tuple(param_names)):  # startswith method accepts tuple
            filteredparams.append(param)
            print(param.Name)  # To check if a parameter in the list is not supposed to be deleted

    # Delete all parameters in the list
    t = Transaction(doc, "Delete parameters")
    t.Start()
    for param in filteredparams:
        doc.Delete(param.Id)
    t.Commit()
