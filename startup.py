# Imports
from pyrevit import HOST_APP
from pyrevit import revit, DB, UI
import os
from Autodesk.Revit.UI import SelectionUIOptions
from System.Windows.Media.Imaging import BitmapImage
import System

from pyrevit import forms
from pyrevit.framework import Threading
from LineNumberPane.pane import MyPane
from LineNumberPane.config import is_visible, set_visible

# Revit Selection Options
uidoc = HOST_APP.uidoc
opts  = SelectionUIOptions.GetSelectionUIOptions()
opts.DragOnSelection = False
opts.SelectLinks     = False
opts.SelectUnderlay  = True

PANE_AVAILABLE = False

try:
    from pyrevit import forms
    from pyrevit.framework import Threading
    from LineNumberPane.pane import MyPane
    from LineNumberPane.config import is_visible, set_visible
    PANE_AVAILABLE = True
except Exception:
    import traceback
    with open(r'C:\temp\startup_debug.txt', 'a') as log:
        log.write("Pane import error:\n{}\n".format(traceback.format_exc()))

# Restore FP Hook button icon on load
FLAG_FILE  = r'C:\temp\fabrication_hook_enabled.txt'
RIBBON_TAB = 'MC Tools'

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
        icon_dir   = os.path.join(
            bundle_dir,
            'MC Tools.tab',
            'Fabrication.panel',
            'SyncData.splitpushbutton',
            'ToggleFPSync.smartbutton'
        )
        icon_path  = os.path.join(icon_dir, 'on.png' if state else 'off.png')

        if not os.path.exists(icon_path):
            with open(r'C:\temp\startup_debug.txt', 'a') as log:
                log.write("Icon not found: {}\n".format(icon_path))
            return

        uri = System.Uri(icon_path, System.UriKind.Absolute)
        img = BitmapImage(uri)

        for tab in ComponentManager.Ribbon.Tabs:
            if tab.Title != RIBBON_TAB:
                continue
            for panel in tab.Panels:
                try:
                    for item in panel.Source.Items:
                        if hasattr(item, 'Text') and item.Text and 'Auto\nSync' in item.Text:
                            item.LargeImage = img
                            item.Image      = img
                            return
                except Exception:
                    pass

    except Exception:
        import traceback
        with open(r'C:\temp\startup_debug.txt', 'w') as log:
            log.write("set_hook_icon error:\n{}\n".format(traceback.format_exc()))

    finally:
        try:
            HOST_APP.uiapp.ViewActivated -= set_hook_icon
        except Exception:
            pass

# Subscribe - will fire once the first view is activated after Revit loads
HOST_APP.uiapp.ViewActivated += set_hook_icon


# dockable pane ===============================================================

if PANE_AVAILABLE:
    def _pane_defer(func):
        try:
            Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                Threading.DispatcherPriority.Background,
                System.Action(func)
            )
        except Exception:
            try:
                func()
            except Exception:
                pass

    def _ensure_my_pane_registered():
        try:
            if not forms.is_registered_dockable_panel(MyPane):
                forms.register_dockable_panel(MyPane, default_visible=False)
        except Exception:
            import traceback
            with open(r'C:\temp\startup_debug.txt', 'a') as log:
                log.write("Pane register error:\n{}\n".format(traceback.format_exc()))

    def _restore_my_pane_visibility():
        try:
            if is_visible():
                forms.open_dockable_panel(MyPane)
            else:
                forms.close_dockable_panel(MyPane)
        except Exception:
            pass

    def _sync_my_pane_visibility():
        try:
            dockable = forms.get_dockable_panel(MyPane)
            set_visible(dockable.IsShown())
        except Exception:
            pass

    def _on_my_pane_visibility_changed(sender, args):
        _pane_defer(_sync_my_pane_visibility)

    def _on_my_pane_document_opened(sender, args):
        try:
            HOST_APP.uiapp.Application.DocumentOpened -= _on_my_pane_document_opened
        except Exception:
            pass

        _pane_defer(_restore_my_pane_visibility)

    _ensure_my_pane_registered()

    try:
        HOST_APP.uiapp.DockableFrameVisibilityChanged += _on_my_pane_visibility_changed
    except Exception:
        pass

    try:
        if HOST_APP.uidoc is not None:
            _pane_defer(_restore_my_pane_visibility)
        else:
            HOST_APP.uiapp.Application.DocumentOpened += _on_my_pane_document_opened
    except Exception:
        try:
            HOST_APP.uiapp.Application.DocumentOpened += _on_my_pane_document_opened
        except Exception:
            pass