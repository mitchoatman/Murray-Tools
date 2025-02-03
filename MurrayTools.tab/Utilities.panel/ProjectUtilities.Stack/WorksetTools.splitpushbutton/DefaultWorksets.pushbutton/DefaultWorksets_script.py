from __future__ import print_function
from Autodesk.Revit.DB import Workset, Transaction, BuiltInCategory, FilteredWorksetCollector
from pyrevit import revit

doc = __revit__.ActiveUIDocument.Document

WorksetNames = []
LevelNames = []

AllWorksets = FilteredWorksetCollector(doc)
for c in AllWorksets:
	WorksetNames.append(c.Name)

WorksetToAdd = []
WorksetToAdd.append('MURRAY Levels and Grids')
WorksetToAdd.append('LINK - ARCH')
WorksetToAdd.append('LINK - STRUCT')
WorksetToAdd.append('LINK - MEP')
WorksetToAdd.append('POINTLAYOUT')

worksetaddedlist = []

if not revit.doc.IsWorkshared and revit.doc.CanEnableWorksharing:
    revit.doc.EnableWorksharing('Workset1','Workset1')

t = Transaction(doc)
t.Start('Create Worksets')
try:
    WorksetList = list(set(WorksetToAdd).difference(set(WorksetNames)))
    if len(WorksetList) > 0:
        for wset in WorksetList:
            Workset.Create(doc, str(wset))
            worksetaddedlist.append(wset)
        print ('Added Workset(s):')
        print (*worksetaddedlist,sep='\n')
    else:
        print ('Worksets already exist')
except:
    pass
t.Commit()

if doc.IsModelInCloud:
    t = Transaction(doc)
    t.Start('Create Worksets')
    try:
        # Enable worksharing
        WorksharingUtils.EnableWorksharing(doc, "Workset1", "Workset1")
        print("Worksharing enabled successfully.")
    except:
        pass
    t.Commit()