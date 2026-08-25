# -*- coding: utf-8 -*-
"""Toggle the Fabrication Auto-Update hook on or off.
Persists state between Revit sessions."""

import os
import clr

from pyrevit import script, forms

# .NET namespaces for native Windows balloon notification
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import Icon, SystemIcons

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


def show_balloon_notification(title, message, icon_path=None, timeout=5000):
    """Displays a native Windows balloon notification."""
    notify_icon = NotifyIcon()
    try:
        if icon_path and os.path.exists(icon_path):
            notify_icon.Icon = Icon(icon_path)
        else:
            notify_icon.Icon = SystemIcons.Information

        notify_icon.Visible = True
        notify_icon.ShowBalloonTip(timeout, title, message, ToolTipIcon.Info)
    except Exception:
        pass


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

        # Optional custom .ico in your button bundle folder
        icon_file = os.path.join(os.path.dirname(__file__), 'Murray.ico')
        if not os.path.exists(icon_file):
            icon_file = None

        show_balloon_notification(
            'Hook {}'.format('Enabled' if new_state else 'Disabled'),
            'Workset Hook is now {}\n\n{}'.format(
                'ON' if new_state else 'OFF',
                'Workset will change when view is changed.'
                if new_state
                else 'Workset will NOT change when view is changed.'
            ),
            icon_path=icon_file,
            timeout=5000
        )


if __name__ == '__main__':
    main()