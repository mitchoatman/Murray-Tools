
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction
from pyrevit import script, forms
import sys

# get the current Revit document
doc = __revit__.ActiveUIDocument.Document

# specify the folder where the files are located
CompleteFolderPath = "C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES"
try:
    # prompt the user to select the family file
    family_path = forms.pick_file(file_ext='rfa', init_dir=str(CompleteFolderPath), multi_file=True, title='Select the families to insert')[0]
    
    t = Transaction(doc, 'Load Family')
    #Start Transaction
    t.Start()
    doc.LoadFamily(str(family_path))
    #End Transaction
    t.Commit()
except:
    sys.exit()