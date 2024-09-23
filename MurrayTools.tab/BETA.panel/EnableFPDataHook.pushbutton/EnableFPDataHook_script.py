import os
import clr
from pyrevit import EXEC_PARAMS, script, forms
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from Autodesk.Revit.DB import Transaction, SynchronizeWithCentralOptions, RelinquishOptions
from SharedParam.Add_Parameters import Shared_Params

Shared_Params()

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Check if the model is workshared
if doc.IsWorkshared:
    # Start a transaction
    t = Transaction(doc, "Sync with Central")
    t.Start()

    # Define relinquishing options (customize as needed)
    relinquishOptions = RelinquishOptions(False)
    relinquishOptions.StandardWorksets = True
    relinquishOptions.ViewWorksets = True
    relinquishOptions.FamilyWorksets = True
    relinquishOptions.UserWorksets = True
    relinquishOptions.CheckedOutElements = True

    # Set up synchronization options
    syncOptions = SynchronizeWithCentralOptions()
    syncOptions.SetRelinquishOptions(relinquishOptions)
    syncOptions.Compact = False  # Set to True if you want to compact the central file

    # Sync with central
    doc.SynchronizeWithCentral(syncOptions)

    # Commit the transaction
    t.Commit()

try:
    my_file = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\Murray Ribbon\Murray.extension\hooks\UpdateFPData_doc-updater.txt'
    base = os.path.splitext(my_file)[0]
    os.rename(my_file, base + '.py')

    if EXEC_PARAMS.executed_from_ui:
        logger = script.get_logger()
        results = script.get_results()
        # re-load pyrevit session.
        logger.info('Reloading....')
        sessionmgr.reload_pyrevit()
        results.newsession = sessioninfo.get_session_uuid()
except:
    my_file = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\Murray Ribbon\Murray.extension\hooks\UpdateFPData_doc-updater.py'
    base = os.path.splitext(my_file)[0]
    os.rename(my_file, base + '.txt')

    if EXEC_PARAMS.executed_from_ui:
        logger = script.get_logger()
        results = script.get_results()
        # re-load pyrevit session.
        logger.info('Reloading....')
        sessionmgr.reload_pyrevit()
        results.newsession = sessioninfo.get_session_uuid()
