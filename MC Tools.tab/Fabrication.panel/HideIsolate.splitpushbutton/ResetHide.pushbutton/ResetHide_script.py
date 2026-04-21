from Autodesk.Revit.DB import *

# .NET Imports
import os, clr
clr.AddReference("System")
from System.Collections.Generic import List

doc   = __revit__.ActiveUIDocument.Document

all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElementIds()
unhide_elements = List[ElementId](all_elements)

t = Transaction(doc, 'UnHide Elements')
t.Start()
doc.ActiveView.UnhideElements(unhide_elements)
t.Commit()