from __future__ import print_function
from Autodesk.Revit.DB import Workset, Transaction, FilteredWorksetCollector
from pyrevit import revit

doc = __revit__.ActiveUIDocument.Document

WorksetToAdd = [
    'MURRAY Levels and Grids',
    'LINKS',
    'POINTLAYOUT'
]

worksetaddedlist = []

# Step 1: Enable worksharing if not already enabled (outside any transaction)
worksharing_enabled = False
if not doc.IsWorkshared:
    try:
        if doc.IsModelInCloud:
            if doc.CanEnableCloudWorksharing():
                doc.EnableCloudWorksharing()
                print("Cloud worksharing enabled successfully.")
                worksharing_enabled = True
            else:
                print("Cloud worksharing cannot be enabled for this model.")
                raise Exception("Worksharing enablement prerequisites not met.")
        elif doc.CanEnableWorksharing():
            doc.EnableWorksharing('Workset1', 'Workset1')
            print("Local worksharing enabled successfully.")
            worksharing_enabled = True
        else:
            print("Worksharing cannot be enabled for this model.")
            raise Exception("Worksharing enablement prerequisites not met.")
    except Exception as e:
        print("Error enabling worksharing: {}".format(str(e)))
        # Halt script execution if worksharing is required
else:
    worksharing_enabled = True
    print("Worksharing is already enabled.")

# Step 2: Verify worksharing is active before proceeding
if not doc.IsWorkshared:
    print("Worksharing enablement failed; cannot create worksets.")
else:
    # Collect current workset names (refreshed after potential enablement)
    WorksetNames = []
    AllWorksets = FilteredWorksetCollector(doc)
    for c in AllWorksets:
        WorksetNames.append(c.Name)

    # Step 3: Create additional worksets if needed
    t = Transaction(doc, 'Create Worksets')
    t.Start()
    try:
        WorksetList = list(set(WorksetToAdd).difference(set(WorksetNames)))
        if len(WorksetList) > 0:
            for wset in WorksetList:
                Workset.Create(doc, str(wset))
                worksetaddedlist.append(wset)
            print('Added Workset(s):')
            print(*worksetaddedlist, sep='\n')
        else:
            print('Specified worksets already exist')
        t.Commit()
    except Exception as e:
        t.RollBack()
        print("Error creating worksets: {}".format(str(e)))