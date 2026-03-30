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

import os
import subprocess
import time
from pyrevit import forms
from pyrevit.framework import Media

def run_batch_file_with_progress(batch_file_path):
    if not os.path.isfile(batch_file_path):
        return False

    try:
        # Start the batch file process
        process = subprocess.Popen(batch_file_path, shell=True)

        # Define colors (ARGB format)
        accent_color = Media.Color.FromArgb(0xFF, 0xFF, 0xFF, 0x00)  # Bright yellow
        text_color   = Media.Color.FromArgb(0xFF, 0x00, 0x00, 0x00)  # Pure black

        # Create the progress bar instance
        pb = forms.ProgressBar(
            title='Running database update batch file...',
            cancellable=True
        )

        # Apply custom accent brush (progress fill/indicator)
        pb.Resources["pyRevitAccentBrush"] = Media.SolidColorBrush(accent_color)

        # Apply black text color to title and progress percentage
        pb.Foreground = Media.SolidColorBrush(text_color)

        with pb:
            counter = 0.0
            max_value = 100.0

            while process.poll() is None:
                if pb.cancelled:
                    process.terminate()
                    return False

                # Update with a cycling pseudo-progress animation
                pb.update_progress(counter % max_value, max_value)
                counter += 0.5  # Adjust increment for smoother/faster animation
                time.sleep(0.3)

        # Process completed normally
        return True

    except Exception as ex:
        forms.alert("Error running batch file: {}".format(str(ex)))
        return False

    except Exception as ex:
        # Consider logging the exception if needed
        forms.alert("Error running batch file: {}".format(str(ex)))
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
