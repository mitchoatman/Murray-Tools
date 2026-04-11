from Autodesk.Revit.DB import Workset, Transaction, FilteredWorksetCollector, WorksetKind
from Autodesk.Revit.UI import TaskDialog

doc = __revit__.ActiveUIDocument.Document

# Desired worksets
WORKSETS_TO_ADD = [
    'MURRAY Levels and Grids',
    'LINKS',
    'POINTLAYOUT'
]


def get_user_workset_names(document):
    """Return a list of existing user workset names."""
    return [ws.Name for ws in FilteredWorksetCollector(document).OfKind(WorksetKind.UserWorkset)]


# --------------------------------------------------------------------
# Step 1: Enable worksharing if needed
# --------------------------------------------------------------------
if not doc.IsWorkshared:
    try:
        if doc.IsModelInCloud:
            if doc.CanEnableCloudWorksharing():
                doc.EnableCloudWorksharing()
                TaskDialog.Show("Success", "Cloud worksharing enabled successfully.")
            else:
                raise Exception("Cloud worksharing cannot be enabled for this model.")
        elif doc.CanEnableWorksharing():
            # First name = shared workset
            # Second name = levels and grids workset
            doc.EnableWorksharing('Workset1', 'MURRAY Levels and Grids')
        else:
            raise Exception("Worksharing cannot be enabled for this model.")

    except Exception as e:
        TaskDialog.Show("Error", "Error enabling worksharing:\n\n{}".format(str(e)))

# --------------------------------------------------------------------
# Step 2: Stop if worksharing still isn't enabled
# --------------------------------------------------------------------
if not doc.IsWorkshared:
    TaskDialog.Show("Error", "Worksharing is not enabled. Cannot create worksets.")

else:
    existing_worksets = get_user_workset_names(doc)
    missing_worksets = [ws for ws in WORKSETS_TO_ADD if ws not in existing_worksets]

    if not missing_worksets:
        TaskDialog.Show("Worksets", "All specified worksets already exist.")

    else:
        t = Transaction(doc, 'Create Murray Worksets')
        t.Start()

        try:
            for workset_name in missing_worksets:
                Workset.Create(doc, workset_name)

            t.Commit()

            TaskDialog.Show(
                "Worksets Created",
                "Added the following workset(s):\n\n{}".format("\n".join(missing_worksets))
            )

        except Exception as e:
            t.RollBack()
            TaskDialog.Show("Error", "Error creating worksets:\n\n{}".format(str(e)))