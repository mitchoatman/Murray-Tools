# -*- coding: UTF-8 -*-
import os

# Import .NET namespaces for native Windows Balloon Notification & Icons
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import Icon

def show_balloon_notification(title, message, icon_path=None, timeout=5000):
    """Displays a native Windows balloon notification using a custom .ico file if available."""
    notify_icon = NotifyIcon()
    try:
        if icon_path and os.path.exists(icon_path):
            notify_icon.Icon = Icon(icon_path)
        else:
            # Fallback to a standard system icon if the custom .ico isn't found
            from System.Drawing import SystemIcons
            notify_icon.Icon = SystemIcons.Information
            
        notify_icon.Visible = True
        notify_icon.ShowBalloonTip(timeout, title, message, ToolTipIcon.Info)
    except Exception:
        pass

# Optional: Setup icon path if you want to use your custom Murray.ico
# path, filename = __file__.rsplit('\\', 1) if '\\' in __file__ else ('', __file__)
# NewFilename = r'\Murray.ico'
# icon_file = path + NewFilename
icon_file = None # Set to your icon path if desired, e.g., r"C:\path\to\Murray.ico"

# Show the native balloon notification
show_balloon_notification(
    "Database Update", 
    """The Database and Support files are being synced to your Hard Drive.
Once sync is complete, you must reload the Fabrication database inside your Revit project.""",
    icon_path=icon_file
)

# Launch the sync process
os.startfile(r"C:\\Egnyte\Shared\\Engineering\\031407MEP\\Sync-Database\\Update Database - notimer.lnk")