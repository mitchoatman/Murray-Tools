import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, IFamilyLoadOptions
from pyrevit import forms
import clr
import sys

# Define a class implementing IFamilyLoadOptions
class FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        # Automatically replace the family
        overwriteParameterValues[0] = True
        return True  # Continue loading the family

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, overwriteParameterValues):
        # Automatically replace the shared family
        overwriteParameterValues[0] = True
        return True  # Continue loading the shared family

# Get the current Revit document
doc = __revit__.ActiveUIDocument.Document

# Specify the folder where the files are located
CompleteFolderPath = r"C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES"

try:
    # Prompt the user to select the family file
    family_path = forms.pick_file(file_ext='rfa', init_dir=str(CompleteFolderPath), multi_file=False, title='Select the family to insert')
    
    # Start a transaction
    t = Transaction(doc, 'Load Family')
    t.Start()
    
    # Load the family with custom FamilyLoadOptions
    load_options = FamilyLoadOptions()
    loaded_family = clr.StrongBox[DB.Family]()  # Create a StrongBox[Family]
    success = doc.LoadFamily(family_path, load_options, loaded_family)
    
    if success:
        print "Family '{}' loaded successfully.".format(family_path)
        if loaded_family.Value:
            print "Loaded Family Name: {}".format(loaded_family.Value.Name)
    else:
        print "Failed to load family '{}'.".format(family_path)
    
    # Commit the transaction
    t.Commit()

except Exception as e:
    print "An error occurred: {}".format(e)
    sys.exit()
