

from pyrevit import forms
import Autodesk
from Autodesk.Revit.DB import *
import System
import os


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)


folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_FavoriteViews.txt')

with open((filepath), 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

print lines[0] == str(file_name)

AllViews = FilteredElementCollector(doc).OfClass(View)

AllViewNames = []

for view in AllViews:
    AllViewNames.append(view.Name)

if lines[0] == str(file_name):
    for view in AllViews:
        ViewToOpen = doc.GetElement(view)
        uidoc.RequestViewChange(ViewToOpen)