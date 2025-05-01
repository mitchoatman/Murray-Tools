from pyrevit import revit, DB, forms
from pyrevit.forms import ProgressBar
import os
import subprocess
import time
from Autodesk.Revit.DB import SynchronizeWithCentralOptions, RelinquishOptions, TransactWithCentralOptions

def sync_with_central():
    try:
        doc = revit.doc
        if not doc.IsWorkshared:
            return True

        relinquishOptions = RelinquishOptions(False)
        relinquishOptions.StandardWorksets = True
        relinquishOptions.ViewWorksets = True
        relinquishOptions.FamilyWorksets = True
        relinquishOptions.UserWorksets = True
        relinquishOptions.CheckedOutElements = True

        syncOptions = SynchronizeWithCentralOptions()
        syncOptions.SetRelinquishOptions(relinquishOptions)
        syncOptions.Compact = False

        transactOptions = TransactWithCentralOptions()
        doc.SynchronizeWithCentral(transactOptions, syncOptions)

        return True

    except:
        pass

def run_batch_file_with_progress(batch_file_path):
    if not os.path.isfile(batch_file_path):
        return False
    try:
        # Start batch file process
        process = subprocess.Popen(batch_file_path, shell=True)

        with ProgressBar(title='Running database update batch file...', cancellable=True) as pb:
            counter = 0
            max_value = 100
            while process.poll() is None:
                if pb.cancelled:
                    process.terminate()
                    return False
                pb.update_progress(counter % max_value, max_value)
                counter += 0.1
                time.sleep(0.3)

        return True

    except Exception as e:
        return False

def reload_fabrication_config():
    try:
        doc = revit.doc
        fab_config = DB.FabricationConfiguration.GetFabricationConfiguration(doc)
        if fab_config is None:
            return False

        with revit.Transaction("Reload Fabrication Configuration"):
            fab_config.ReloadConfiguration()
        return True
        forms.show_balloon('Reload Config', 'Fabrication config reloaded')
    except Exception as e:
        return False

if __name__ == "__main__":
    batch_file_path = r"C:\Egnyte\Shared\Engineering\031407MEP\Sync-Database\Sync-Database to Local - notimer.bat"
    sync_with_central()
    if run_batch_file_with_progress(batch_file_path):
        reload_fabrication_config()
