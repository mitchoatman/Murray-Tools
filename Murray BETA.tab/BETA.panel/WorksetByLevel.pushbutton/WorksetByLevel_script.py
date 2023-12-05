from __future__ import print_function
from Autodesk.Revit.DB import Workset, Transaction, FilteredElementCollector, BuiltInCategory, FilteredWorksetCollector

doc = __revit__.ActiveUIDocument.Document

WorksetNames = []
LevelNames = []

AllWorksets = FilteredWorksetCollector(doc)
for c in AllWorksets:
	WorksetNames.append(c.Name)

AllLevels =  FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

worksetaddedlist = []
t = Transaction(doc)
t.Start('Create Level Worksets')
try:
    for Level in AllLevels:
        LevelNames.append(Level.Name)
    WorksetList = list(set(LevelNames).difference(set(WorksetNames)))
    if len(WorksetList) > 0:
        for wset in WorksetList:
            Workset.Create(doc, str(wset))
            worksetaddedlist.append(wset)
        print ('Added Workset(s):')
        print (*worksetaddedlist,sep='\n')
    else:
        print ('Worksets for every level already exist')
except:
    pass
t.Commit()