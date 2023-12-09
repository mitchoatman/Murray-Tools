
#Imports
import Autodesk
from pyrevit import revit
from Autodesk.Revit.DB import Transaction, BuiltInParameter
from Autodesk.Revit.UI import Selection
import sys

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Get the active workset
active_workset = doc.GetWorksetTable().GetActiveWorksetId()
active_workset_name = doc.GetWorksetTable().GetWorkset(active_workset).Name

try:
    #this is start of select element(s)
    sel = uidoc.Selection.PickObjects(Selection.ObjectType.Element, 'Select Elements')
    selected_element = [doc.GetElement( elId ) for elId in sel]
    #this is end of select element(s)

    worksetint = active_workset.IntegerValue


    # START TRANSACTION
    t = Transaction(doc, 'Set Workset for Elements')
    t.Start()
    for el in selected_element:
        workset_param = el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        workset_param.Set(worksetint)
    t.Commit()
except:
    sys.exit()
    