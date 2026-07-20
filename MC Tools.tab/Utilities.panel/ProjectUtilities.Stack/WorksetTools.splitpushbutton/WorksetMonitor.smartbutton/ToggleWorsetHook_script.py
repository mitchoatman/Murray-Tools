# -*- coding: utf-8 -*-
"""Toggle the Fabrication Auto-Update hook on or off.
Persists state between Revit sessions."""

import os
from pyrevit import script, forms

# ── Config ────────────────────────────────────────────────────────────────────
FLAG_FILE = r'C:\temp\Ribbon_Workset-hook-status.txt'

def get_state():
    try:
        if os.path.exists(FLAG_FILE):
            with open(FLAG_FILE, 'r') as f:
                return f.read().strip().lower() == 'true'
    except Exception:
        pass
    return True  # Default state (enabled)

def set_state(value):
    try:
        folder = os.path.dirname(FLAG_FILE)
        if not os.path.exists(folder):
            os.makedirs(folder)
        with open(FLAG_FILE, 'w') as f:
            f.write('true' if value else 'false')
        return True
    except Exception as e:
        forms.alert('Could not save state:\n{}'.format(str(e)), title='Error')
        return False

def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    try:
        state = get_state()
        icon_name = 'on.png' if state else 'off.png'
        icon_path = script_cmp.get_bundle_file(icon_name)

        if icon_path and os.path.exists(icon_path):
            ui_button_cmp.set_icon(icon_path)
        else:
            fallback = script_cmp.get_bundle_file('off.png')
            if fallback and os.path.exists(fallback):
                ui_button_cmp.set_icon(fallback)

        return True
    except Exception:
        return True

def main():
    current_state = get_state()
    new_state = not current_state
    
    if set_state(new_state):
        script.toggle_icon(new_state)
        forms.show_balloon(
            'Hook {}'.format('Enabled' if new_state else 'Disabled'),
            'Workset Hook is now {}\n\n{}'.format(
                'ON' if new_state else 'OFF',
                'Workset will change when view is changed.' 
                if new_state 
                else 'Workset will NOT change when view is changed.'
            )
        )

if __name__ == '__main__':
    main()