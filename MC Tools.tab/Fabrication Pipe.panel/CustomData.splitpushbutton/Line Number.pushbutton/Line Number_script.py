# -*- coding: utf-8 -*-
"""Toggle Line Number dockable pane."""
from pyrevit import forms
from pyrevit.revit import ui
from pyrevit.framework import Threading, System
import pyrevit.extensions as exts
from pyrevit.coreutils.ribbon import ICON_MEDIUM

from LineNumberPane.pane import MyPane
from LineNumberPane.config import set_visible


def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    def set_icon(is_shown):
        icon_name = exts.DEFAULT_ON_ICON_FILE if is_shown else exts.DEFAULT_OFF_ICON_FILE
        icon = ui.resolve_icon_file(script_cmp.directory, icon_name)
        ui_button_cmp.set_icon(icon, icon_size=ICON_MEDIUM)

    def refresh_icon():
        try:
            if forms.is_registered_dockable_panel(MyPane):
                dockable = forms.get_dockable_panel(MyPane)
                set_icon(dockable.IsShown())
            else:
                set_icon(False)
        except Exception:
            set_icon(False)

    def on_visibility_changed(sender, args):
        def deferred():
            try:
                if forms.is_registered_dockable_panel(MyPane):
                    dockable = forms.get_dockable_panel(MyPane)
                    shown = dockable.IsShown()
                    set_visible(shown)
                    set_icon(shown)
                else:
                    set_icon(False)
            except Exception:
                set_icon(False)

        Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
            Threading.DispatcherPriority.Background,
            System.Action(deferred)
        )

    def on_document_opened(sender, args):
        try:
            __rvt__.Application.DocumentOpened -= on_document_opened
        except Exception:
            pass
        refresh_icon()

    set_icon(False)

    try:
        __rvt__.DockableFrameVisibilityChanged += on_visibility_changed
    except Exception:
        pass

    try:
        if __rvt__.ActiveUIDocument is not None:
            refresh_icon()
        else:
            __rvt__.Application.DocumentOpened += on_document_opened
    except Exception:
        pass

    return True


if __name__ == "__main__":
    if not forms.is_registered_dockable_panel(MyPane):
        forms.register_dockable_panel(MyPane, default_visible=False)

    dockable = forms.get_dockable_panel(MyPane)
    new_state = not dockable.IsShown()
    forms.toggle_dockable_panel(MyPane, new_state)
    set_visible(new_state)