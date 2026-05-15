# Imports
from pyrevit import HOST_APP
from pyrevit import revit, DB, UI
import os
from Autodesk.Revit.UI import SelectionUIOptions
from System.Windows.Media.Imaging import BitmapImage
import System

# Revit Selection Options 
uidoc = HOST_APP.uidoc
opts  = SelectionUIOptions.GetSelectionUIOptions()
opts.DragOnSelection = False
opts.SelectLinks     = False
opts.SelectUnderlay  = True

# Restore FP Hook button icon on load 
FLAG_FILE  = r'C:\temp\fabrication_hook_enabled.txt'
RIBBON_TAB = 'MurrayTools'

def get_hook_state():
    try:
        if os.path.exists(FLAG_FILE):
            with open(FLAG_FILE, 'r') as f:
                state = f.read().strip().lower()
                return state == 'true'
    except Exception:
        pass
    return False

from Autodesk.Revit.UI.Events import ViewActivatedEventArgs

def set_hook_icon(sender, args):
    """Fires on first view activated - ribbon is fully built by then."""
    try:
        import clr
        clr.AddReference('AdWindows')
        from Autodesk.Windows import ComponentManager
        from System.Windows.Media.Imaging import BitmapImage
        import System

        state      = get_hook_state()
        bundle_dir = os.path.dirname(os.path.abspath(__file__))
        icon_dir   = os.path.join(bundle_dir, 'MurrayTools.tab', 'Fabrication.panel', 'ToggleFPSync.pushbutton')
        icon_path  = os.path.join(icon_dir, 'on.png' if state else 'off.png')

        uri = System.Uri(icon_path, System.UriKind.Absolute)
        img = BitmapImage(uri)

        for tab in ComponentManager.Ribbon.Tabs:
            if 'Murray' not in tab.Title:
                continue
            for panel in tab.Panels:
                try:
                    for item in panel.Source.Items:
                        if hasattr(item, 'Text') and item.Text and 'Auto\nSync' in item.Text:
                            item.LargeImage = img
                            item.Image      = img
                except Exception:
                    pass

    except Exception as e:
        import traceback
        with open(r'C:\temp\startup_debug.txt', 'w') as log:
            log.write("set_hook_icon error:\n{}\n".format(traceback.format_exc()))

    finally:
        # Unsubscribe after first run - only need to set icon once
        HOST_APP.uiapp.ViewActivated -= set_hook_icon

# Subscribe - will fire once the first view is activated after Revit loads
HOST_APP.uiapp.ViewActivated += set_hook_icon


# test dockable panel =========================================================

# from pyrevit import forms
# import os.path as op

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