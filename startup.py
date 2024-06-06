# Imports

from pyrevit import HOST_APP, framework
from pyrevit import DB
from pyrevit import forms

# Variables
uidoc = HOST_APP.uidoc


# Get And Modify Options
from Autodesk.Revit.UI import SelectionUIOptions

opts = SelectionUIOptions.GetSelectionUIOptions()
opts.DragOnSelection  = False
opts.SelectLinks	= False
opts.SelectUnderlay = True

#Other Options: (Select True/False)
#opts.SelectFaces	= False # / True
#opts.SelectPinned	= False # / True
#opts.ActivateControlsAndDimensionsOnMultiSelect  = False #/ True