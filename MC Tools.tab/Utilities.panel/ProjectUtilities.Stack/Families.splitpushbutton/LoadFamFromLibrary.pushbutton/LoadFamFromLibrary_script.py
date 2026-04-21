import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, IFamilyLoadOptions
from pyrevit import forms
import clr
import sys

# Define a class implementing IFamilyLoadOptions
class FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues[0] = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, overwriteParameterValues):
        overwriteParameterValues[0] = True
        return True

# Get the current Revit document
doc = __revit__.ActiveUIDocument.Document

# Specify the folder where the files are located
CompleteFolderPath = r"C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES"

# Initialize transaction variable
t = None

try:
    # Prompt the user to select the family file
    family_path = forms.pick_file(file_ext='rfa', init_dir=str(CompleteFolderPath), multi_file=False, title='Select the family to insert')
    
    # Check if user canceled the dialog
    if not family_path:
        print "Operation canceled by user."
        sys.exit(0)
    
    # Start a transaction
    t = Transaction(doc, 'Load Family')
    t.Start()
    
    # Load the family with custom FamilyLoadOptions
    load_options = FamilyLoadOptions()
    loaded_family = clr.StrongBox[DB.Family]()
    success = doc.LoadFamily(family_path, load_options, loaded_family)
    
    if success and loaded_family.Value:
        forms.show_balloon('Load Family', "Family '{}' loaded successfully.".format(family_path))
        forms.show_balloon('Load Family', "Loaded Family Name: {}".format(loaded_family.Value.Name))
    else:
        print "Failed to load family '{}'.".format(family_path)
    
    # Commit the transaction
    t.Commit()

except Exception as e:
    print "An error occurred: {}".format(e)
    # Roll back the transaction if it was started
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()
    sys.exit(1)

finally:
    # Ensure transaction is disposed if it exists
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()
    if t:
        t.Dispose()