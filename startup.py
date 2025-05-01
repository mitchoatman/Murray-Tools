# Imports

from pyrevit import HOST_APP, framework
from pyrevit import revit, DB, UI
from pyrevit import forms, routes
import sys
import time
import os.path as op

# test dockable panel =========================================================

# class DockableExample(forms.WPFPanel):
    # panel_title = "pyRevit Dockable Panel Title"
    # panel_id = "3110e336-f81c-4927-87da-4e0d30d4d64a"
    # panel_source = op.join(op.dirname(__file__), "DockableExample.xaml")

    # def do_something(self, sender, args):
        # forms.alert("Voila!!!")


# if not forms.is_registered_dockable_panel(DockableExample):
    # forms.register_dockable_panel(DockableExample)
# else:
    # print("Skipped registering dockable pane. Already exists.")


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