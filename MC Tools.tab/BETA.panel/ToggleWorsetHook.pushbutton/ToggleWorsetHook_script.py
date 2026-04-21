# -*- coding: utf-8 -*-
"""Toggle the Fabrication Auto-Update hook on or off.
Persists state between Revit sessions."""

import os
from pyrevit import forms, revit
from System.Windows.Media.Imaging import BitmapImage
import System

# ── Config ────────────────────────────────────────────────────────────────────
FLAG_FILE = r'C:\temp\workset_hook_enabled.txt'
RIBBON_TAB = 'MurrayTools'  # Change to match your ribbon tab exactly

def get_state():
    try:
        if os.path.exists(FLAG_FILE):
            with open(FLAG_FILE, 'r') as f:
                state = f.read().strip()
                return state.lower() == 'true'
    except Exception:
        pass
    return True

def set_state(value):
    try:
        if not os.path.exists(r'C:\temp'):
            os.makedirs(r'C:\temp')
        with open(FLAG_FILE, 'w') as f:
            f.write('true' if value else 'false')
        return True
    except Exception as e:
        forms.alert('Could not save state:\n{}'.format(str(e)), title='Error')
        return False

def swap_icon(new_state):
    try:
        from pyrevit import HOST_APP
        bundle_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path  = os.path.join(bundle_dir, 'ws_on.png' if new_state else 'ws_off.png')
        if not os.path.exists(icon_path):
            print("Icon not found: {}".format(icon_path))
            return
        uri = System.Uri(icon_path, System.UriKind.Absolute)
        img = BitmapImage(uri)
        for panel in HOST_APP.uiapp.GetRibbonPanels(RIBBON_TAB):
            for item in panel.GetItems():
                if 'ToggleWorsetHook' in item.ItemText:
                    item.Image      = img
                    item.LargeImage = img
                    return
        print("Icon swap: button 'Enable Workset' not found on tab '{}'".format(RIBBON_TAB))
    except Exception as e:
        print("Icon swap error: {}".format(str(e)))

# ── Toggle ────────────────────────────────────────────────────────────────────
current_state = get_state()
new_state     = not current_state

if set_state(new_state):
    swap_icon(new_state)
    forms.show_balloon(
        'Hook {}'.format('Enabled' if new_state else 'Disabled'),
        'Workset Hook {}\n\n{}'.format(
            'ON' if new_state else 'OFF',
            'Workset will change when view is changed.'
            if new_state else
            'Workset will NOT change when view is changed.'
        )
    )